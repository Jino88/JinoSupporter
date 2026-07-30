using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace InferenceDataAIService.Wpf;

// This is deliberately separate from the WPF surface.  It consumes the
// persisted numeric capture and only starts rendering after its materialized
// catalog has passed all grouping invariants.
internal static class TwoLevelGroupAnalysisEngine
{
    internal const string PlanSchemaVersion = "two-level-group-plan-v1";
    private const string ValidationSchemaVersion = "two-level-group-validation-v1";
    private const string RefinementSchemaVersion = "group-refinement-v1";
    private const int DefaultClustersPerRequest = 40;
    private const int MaximumRefinementAttempts = 2;
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false);
    private static readonly Regex StableSlug = new("^[a-z0-9]+(?:[._-][a-z0-9]+)*$", RegexOptions.Compiled);
    private static readonly IReadOnlyDictionary<string, string> RendererForTopLevel = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["acoustic-ng-dashboard"] = "acoustic-dashboard-v1",
        ["quality-ng-dashboard"] = "quality-dashboard-v1",
        ["measurement-dimension-dashboard"] = "measurement-dashboard-v1",
        ["function-process-dashboard"] = "function-process-dashboard-v1",
        ["tension-dashboard"] = "tension-dashboard-v1",
        ["general-table-dashboard"] = "general-table-dashboard-v1"
    };

    internal static bool HasValidatedCatalog(string batchDirectory)
    {
        var plan = Path.Combine(batchDirectory, "group-plan.json");
        var catalog = Path.Combine(batchDirectory, "group-catalog.json");
        var validation = Path.Combine(batchDirectory, "group-validation-report.json");
        if (!File.Exists(plan) || !File.Exists(catalog) || !File.Exists(validation)) return false;
        try
        {
            using var planDocument = JsonDocument.Parse(File.ReadAllText(plan, Encoding.UTF8));
            using var validationDocument = JsonDocument.Parse(File.ReadAllText(validation, Encoding.UTF8));
            return string.Equals(planDocument.RootElement.GetProperty("schemaVersion").GetString(), PlanSchemaVersion, StringComparison.Ordinal) &&
                   validationDocument.RootElement.TryGetProperty("isValid", out var valid) && valid.GetBoolean();
        }
        catch (JsonException) { return false; }
        catch (KeyNotFoundException) { return false; }
    }

    internal static async Task<TwoLevelGroupAnalysisRunResult> RunAsync(
        TwoLevelGroupAnalysisRequest request,
        Action<string>? log = null,
        CancellationToken cancellationToken = default)
    {
        var batchDirectory = Path.Combine(
            AppRuntimePaths.Current.BatchRootDirectory,
            request.BatchId);
        var database = Path.Combine(batchDirectory, "numeric-capture.sqlite");
        var classification = Path.Combine(batchDirectory, "classification.csv");
        var summary = Path.Combine(batchDirectory, "summary.json");
        if (!Directory.Exists(batchDirectory)) throw new DirectoryNotFoundException($"Structure batch was not found: {batchDirectory}");
        if (!File.Exists(database)) throw new FileNotFoundException("AI grouping reuses the existing numeric capture database; capture it before grouping.", database);
        if (!File.Exists(classification)) throw new FileNotFoundException("Structure classification is required before AI grouping.", classification);
        if (!File.Exists(summary)) throw new FileNotFoundException("Structure summary is required before AI grouping.", summary);

        log?.Invoke("Two-level grouping: reusing the existing numeric capture; rebuilding compact inventory only.");
        var inventory = DocumentInventoryEngine.Write(batchDirectory);
        var evidenceByCluster = ReadClusterEvidence(inventory.ClusterSummaryPath);
        var categoryByCluster = ReadSemanticCategories(inventory.SemanticSummaryPath);
        var clusterByPath = ReadLayoutClusters(inventory.ClusterPath);
        var documentsByPath = ReadInventory(inventory.InventoryPath);
        var scannedPaths = ReadClassificationPaths(classification);
        var declaredScannedCount = ReadScannedCount(summary);
        if (declaredScannedCount != scannedPaths.Count)
            throw new InvalidOperationException($"Structure summary and classification disagree on scanned files: summary={declaredScannedCount}, csv={scannedPaths.Count}.");

        var scanned = new List<TrackedWorkbook>();
        foreach (var path in scannedPaths)
        {
            if (!documentsByPath.TryGetValue(path, out var document))
                throw new InvalidOperationException($"The persisted capture has no inventory record for scanned workbook '{path}'.");
            if (!clusterByPath.TryGetValue(path, out var clusterId))
                throw new InvalidOperationException($"The persisted inventory has no layout cluster for scanned workbook '{path}'.");
            if (!categoryByCluster.TryGetValue(clusterId, out var category))
                throw new InvalidOperationException($"The persisted inventory has no semantic category for layout cluster '{clusterId}'.");
            scanned.Add(new TrackedWorkbook(path, clusterId, category.Category, category.TopLevelCategory, document.RoutingState));
        }

        var invalidRouting = scanned.Where(file => file.RoutingState is not "ASSIGNABLE" and not "EMPTY_LAYOUT" and not "CAPTURE_INCOMPLETE").ToList();
        if (invalidRouting.Count > 0)
            throw new InvalidOperationException($"Scanned workbooks have unsupported routing states: {string.Join(", ", invalidRouting.Take(5).Select(file => $"{file.RelativePath}={file.RoutingState}"))}.");

        var assignableClusters = scanned.Where(file => file.RoutingState == "ASSIGNABLE")
            .Select(file => file.LayoutClusterId).Distinct(StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);
        var missingEvidence = assignableClusters.Where(cluster => !evidenceByCluster.ContainsKey(cluster)).ToList();
        if (missingEvidence.Count > 0)
            throw new InvalidOperationException($"AI grouping input is missing compact evidence for {missingEvidence.Count} assignable layout clusters.");

        var workItems = BuildWorkItems(assignableClusters, categoryByCluster, evidenceByCluster, ClustersPerRequest());
        var refinementDirectory = Path.Combine(batchDirectory, "group-refinement");
        Directory.CreateDirectory(refinementDirectory);
        var allGroups = new List<SecondLevelGroup>();
        var knownByTopLevel = new Dictionary<string, Dictionary<string, SecondLevelGroup>>(StringComparer.Ordinal);
        var localValidationReports = new List<string>();
        foreach (var workItem in workItems)
        {
            cancellationToken.ThrowIfCancellationRequested();
            knownByTopLevel.TryGetValue(workItem.TopLevelCategory, out var known);
            known ??= new Dictionary<string, SecondLevelGroup>(StringComparer.OrdinalIgnoreCase);
            var result = await RefineBatchWithRetryAsync(request.ServiceDirectory, refinementDirectory, workItem, known.Values.ToList(), log, cancellationToken);
            localValidationReports.Add(result.ValidationPath);
            foreach (var group in result.Groups)
            {
                if (known.TryGetValue(group.Id, out var existing))
                {
                    // A reused id is an explicit statement that these compact batches
                    // describe the same second-level form.  Its recipe and renderer
                    // must remain identical before their members can be merged.
                    if (!string.Equals(existing.RendererKey, group.RendererKey, StringComparison.Ordinal) ||
                        !SequenceEqual(existing.ComponentRecipe, group.ComponentRecipe))
                        throw new InvalidOperationException($"Second-level group '{group.Id}' is incompatible across compact batches.");
                    existing.Merge(group);
                }
                else
                {
                    known.Add(group.Id, group);
                    allGroups.Add(group);
                }
            }
            knownByTopLevel[workItem.TopLevelCategory] = known;
        }

        // Scanner-proven empty/incomplete layouts are not normal AI inputs. Keep
        // them visible in named structural groups instead of creating one large
        // fallback bucket that would violate the 1.5% fallback invariant.
        var exceptionGroups = BuildExceptionGroups(scanned);
        var validation = ValidateMergedGroups(scanned, allGroups, evidenceByCluster);
        var validationPath = Path.Combine(batchDirectory, "group-validation-report.json");
        var validationJson = new JsonObject
        {
            ["schemaVersion"] = ValidationSchemaVersion,
            ["generatedAt"] = DateTimeOffset.UtcNow.ToString("O"),
            ["isValid"] = validation.Errors.Count == 0,
            ["counts"] = new JsonObject
            {
                ["scanned"] = scanned.Count,
                ["assignable"] = scanned.Count(file => file.RoutingState == "ASSIGNABLE"),
                ["fallback"] = validation.FallbackCount,
                ["secondLevelGroups"] = allGroups.Count,
                ["namedExceptionGroups"] = exceptionGroups.Count,
                ["compactBatches"] = workItems.Count
            },
            ["errors"] = new JsonArray(validation.Errors.Select(error => (JsonNode?)error).ToArray()),
            ["localValidationReports"] = new JsonArray(localValidationReports.Select(path => (JsonNode?)Path.GetRelativePath(batchDirectory, path).Replace('\\', '/')).ToArray())
        };
        AtomicWrite(validationPath, validationJson.ToJsonString(Indented) + "\n");
        if (validation.Errors.Count > 0)
            throw new InvalidOperationException($"Two-level group validation failed. Rendering was not started. See {validationPath}.");

        var plan = BuildPlan(request.BatchId, scanned, allGroups, exceptionGroups, evidenceByCluster, validation.FallbackCount);
        var planPath = Path.Combine(batchDirectory, "group-plan.json");
        AtomicWrite(planPath, plan.ToJsonString(Indented) + "\n");
        var catalogPath = Path.Combine(batchDirectory, "group-catalog.json");
        AtomicWrite(catalogPath, BuildCatalog(plan).ToJsonString(Indented) + "\n");
        log?.Invoke($"Two-level grouping validated: {allGroups.Count} second-level groups, {exceptionGroups.Count} named exception groups, {scanned.Count} scanned files, fallback {validation.FallbackCount}.");
        return new TwoLevelGroupAnalysisRunResult(batchDirectory, planPath, catalogPath, validationPath, allGroups.Count + exceptionGroups.Count, scanned.Count, validation.FallbackCount);
    }

    private static async Task<RefinementBatchResult> RefineBatchWithRetryAsync(
        string serviceDirectory,
        string refinementDirectory,
        RefinementWorkItem workItem,
        IReadOnlyList<SecondLevelGroup> knownGroups,
        Action<string>? log,
        CancellationToken cancellationToken)
    {
        var errors = new List<string>();
        for (var attempt = 1; attempt <= MaximumRefinementAttempts; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var stem = $"{SafeFilePart(workItem.TopLevelCategory)}-{workItem.Ordinal:D3}-attempt-{attempt}";
            var responsePath = Path.Combine(refinementDirectory, stem + ".json");
            log?.Invoke($"Two-level grouping: {workItem.TopLevelCategory} compact batch {workItem.Ordinal} ({workItem.Clusters.Count} clusters), attempt {attempt}/{MaximumRefinementAttempts}.");
            await InvokeCodexAsync(serviceDirectory, responsePath, BuildPrompt(workItem, knownGroups, errors), log, cancellationToken);
            var parsed = ParseRefinementResponse(responsePath, workItem.TopLevelCategory, workItem.Clusters);
            errors = ValidateRefinementBatch(parsed, workItem, knownGroups).ToList();
            var reportPath = Path.Combine(refinementDirectory, stem + ".validation.json");
            AtomicWrite(reportPath, new JsonObject
            {
                ["schemaVersion"] = "group-refinement-validation-v1",
                ["topLevelCategory"] = workItem.TopLevelCategory,
                ["ordinal"] = workItem.Ordinal,
                ["attempt"] = attempt,
                ["isValid"] = errors.Count == 0,
                ["expectedClusterIds"] = new JsonArray(workItem.Clusters.Select(cluster => (JsonNode?)cluster.LayoutClusterId).ToArray()),
                ["errors"] = new JsonArray(errors.Select(error => (JsonNode?)error).ToArray())
            }.ToJsonString(Indented) + "\n");
            if (errors.Count == 0) return new RefinementBatchResult(parsed.Groups, reportPath);
            log?.Invoke($"Two-level grouping: compact batch {workItem.Ordinal} validation failed ({errors.Count} issues); retrying only this batch.");
        }
        throw new InvalidOperationException($"Two-level grouping could not validate {workItem.TopLevelCategory} compact batch {workItem.Ordinal} after {MaximumRefinementAttempts} attempts. See group-refinement validation reports.");
    }

    private static IReadOnlyList<string> ValidateRefinementBatch(RefinementResponse response, RefinementWorkItem workItem, IReadOnlyList<SecondLevelGroup> knownGroups)
    {
        var errors = new List<string>();
        if (!string.Equals(response.TopLevelCategory, workItem.TopLevelCategory, StringComparison.Ordinal))
            errors.Add($"Response topLevelCategory must be '{workItem.TopLevelCategory}'.");
        var expected = workItem.Clusters.Select(cluster => cluster.LayoutClusterId).ToHashSet(StringComparer.Ordinal);
        var supplied = response.Groups.SelectMany(group => group.LayoutClusterIds).ToList();
        var unknown = supplied.Where(id => !expected.Contains(id)).Distinct(StringComparer.Ordinal).ToList();
        var duplicate = supplied.GroupBy(id => id, StringComparer.Ordinal).Where(group => group.Count() > 1).Select(group => group.Key).ToList();
        var missing = expected.Except(supplied, StringComparer.Ordinal).ToList();
        if (unknown.Count > 0) errors.Add($"Response assigns unknown clusters: {string.Join(", ", unknown.Take(6))}.");
        if (duplicate.Count > 0) errors.Add($"Response assigns a cluster more than once: {string.Join(", ", duplicate.Take(6))}.");
        if (missing.Count > 0) errors.Add($"Response leaves clusters unassigned: {string.Join(", ", missing.Take(6))}.");
        var knownById = knownGroups.ToDictionary(group => group.Id, StringComparer.OrdinalIgnoreCase);
        var duplicateGroupIds = response.Groups.GroupBy(group => group.Id, StringComparer.OrdinalIgnoreCase).Where(group => group.Count() > 1).Select(group => group.Key).ToList();
        if (duplicateGroupIds.Count > 0) errors.Add($"Response repeats second-level group ids: {string.Join(", ", duplicateGroupIds.Take(6))}.");
        foreach (var group in response.Groups)
        {
            if (!StableSlug.IsMatch(group.Id)) errors.Add($"Group id '{group.Id}' is not a stable slug.");
            if (group.ComponentRecipe.Count == 0) errors.Add($"Group '{group.Id}' has no component recipe.");
            if (!GroupCatalogRendererEngine.IsSupportedComponentRecipe(group.ComponentRecipe)) errors.Add($"Group '{group.Id}' uses a component recipe that the captured-component renderer cannot materialize.");
            if (!GroupCatalogRendererEngine.IsRegisteredRendererKey(group.RendererKey)) errors.Add($"Group '{group.Id}' uses unregistered renderer '{group.RendererKey}'.");
            if (!RendererForTopLevel.TryGetValue(workItem.TopLevelCategory, out var expectedRenderer) || !string.Equals(group.RendererKey, expectedRenderer, StringComparison.Ordinal))
                errors.Add($"Group '{group.Id}' must use '{expectedRenderer ?? "a registered renderer"}' for {workItem.TopLevelCategory}.");
            var assigned = group.LayoutClusterIds.Select(id => workItem.Clusters.SingleOrDefault(cluster => cluster.LayoutClusterId == id)).Where(cluster => cluster is not null).Cast<ClusterEvidence>().ToList();
            var recipes = assigned.Select(ComponentRecipeFor).Distinct(StringComparer.Ordinal).ToList();
            if (recipes.Count != 1 || !string.Equals(recipes.SingleOrDefault(), string.Join("|", group.ComponentRecipe), StringComparison.Ordinal))
                errors.Add($"Group '{group.Id}' mixes a component recipe or does not declare its assigned cluster recipe exactly.");
            if (knownById.TryGetValue(group.Id, out var known) &&
                (!string.Equals(known.RendererKey, group.RendererKey, StringComparison.Ordinal) || !SequenceEqual(known.ComponentRecipe, group.ComponentRecipe)))
                errors.Add($"Group '{group.Id}' conflicts with an earlier compact batch; use that id only for the same renderer and component recipe.");
        }
        return errors;
    }

    private static MergedValidation ValidateMergedGroups(IReadOnlyList<TrackedWorkbook> scanned, IReadOnlyList<SecondLevelGroup> groups, IReadOnlyDictionary<string, ClusterEvidence> evidenceByCluster)
    {
        var errors = new List<string>();
        var clusterOwners = new Dictionary<string, SecondLevelGroup>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            if (!GroupCatalogRendererEngine.IsRegisteredRendererKey(group.RendererKey)) errors.Add($"Group '{group.Id}' uses an unregistered renderer.");
            if (!GroupCatalogRendererEngine.IsSupportedComponentRecipe(group.ComponentRecipe)) errors.Add($"Group '{group.Id}' uses a component recipe that the captured-component renderer cannot materialize.");
            if (!RendererForTopLevel.TryGetValue(group.TopLevelCategory, out var expectedRenderer) || !string.Equals(group.RendererKey, expectedRenderer, StringComparison.Ordinal))
                errors.Add($"Group '{group.Id}' renderer is incompatible with top-level category '{group.TopLevelCategory}'.");
            foreach (var clusterId in group.LayoutClusterIds)
            {
                if (!evidenceByCluster.TryGetValue(clusterId, out var evidence)) { errors.Add($"Group '{group.Id}' references unavailable cluster '{clusterId}'."); continue; }
                if (!string.Equals(ComponentRecipeFor(evidence), string.Join("|", group.ComponentRecipe), StringComparison.Ordinal))
                    errors.Add($"Group '{group.Id}' has incompatible component recipe for cluster '{clusterId}'.");
                if (!clusterOwners.TryAdd(clusterId, group)) errors.Add($"Cluster '{clusterId}' overlaps groups '{clusterOwners[clusterId].Id}' and '{group.Id}'.");
            }
        }
        var assignable = scanned.Where(file => file.RoutingState == "ASSIGNABLE").ToList();
        var missing = assignable.Where(file => !clusterOwners.ContainsKey(file.LayoutClusterId)).ToList();
        if (missing.Count > 0) errors.Add($"{missing.Count} ASSIGNABLE workbooks are not assigned to a second-level group.");
        // Scanner exceptions are materialized as named structural groups. The
        // reserved fallback remains empty for this scanned batch and is available
        // only for a future scanner-proven exception that is not otherwise routed.
        const int fallback = 0;
        if (scanned.Count == 0 || fallback / (double)scanned.Count > 0.015d)
            errors.Add($"Fallback is {fallback}/{scanned.Count}; the allowed maximum is 1.5% of scanned workbooks.");
        // Renderer profiles are deliberately category-specific.  Therefore groups
        // sharing a renderer must share its top-level category; each group already
        // has an exact recipe match against every member cluster above.
        foreach (var rendererGroup in groups.GroupBy(group => group.RendererKey, StringComparer.Ordinal))
            if (rendererGroup.Select(group => group.TopLevelCategory).Distinct(StringComparer.Ordinal).Count() != 1)
                errors.Add($"Renderer '{rendererGroup.Key}' is shared by incompatible top-level component recipes.");
        return new MergedValidation(errors, fallback);
    }

    private static JsonObject BuildPlan(string batchId, IReadOnlyList<TrackedWorkbook> scanned, IReadOnlyList<SecondLevelGroup> groups, IReadOnlyList<SecondLevelGroup> exceptionGroups, IReadOnlyDictionary<string, ClusterEvidence> evidenceByCluster, int fallbackCount)
    {
        var normalClusterOwner = groups.SelectMany(group => group.LayoutClusterIds.Select(cluster => (Cluster: cluster, Group: group)))
            .ToDictionary(item => item.Cluster, item => item.Group, StringComparer.Ordinal);
        var exceptionOwner = exceptionGroups.SelectMany(group => group.LayoutClusterIds.Select(cluster => (Cluster: cluster, Group: group)))
            .ToDictionary(item => item.Cluster, item => item.Group, StringComparer.Ordinal);
        var jsonGroups = new JsonArray(groups.OrderBy(group => group.TopLevelCategory, StringComparer.Ordinal).ThenBy(group => group.Id, StringComparer.Ordinal).Select(group => GroupJson(group, evidenceByCluster)).ToArray());
        foreach (var group in exceptionGroups.OrderBy(group => group.Id, StringComparer.Ordinal)) jsonGroups.Add(GroupJson(group, evidenceByCluster));
        jsonGroups.Add(new JsonObject
        {
            ["id"] = "needs-review-fallback",
            ["name"] = "Scanner-proven empty or incomplete capture",
            ["rendererKey"] = "renderNeedsReview",
            ["selector"] = new JsonObject { ["layoutClusterIds"] = new JsonArray(), ["fallback"] = true },
            ["componentRecipe"] = new JsonArray("EMPTY_LAYOUT_OR_CAPTURE_INCOMPLETE"),
            ["representativeFiles"] = new JsonArray(),
            ["importantData"] = new JsonArray("Only scanner-proven exceptions are allowed in this group."),
            ["extractionRule"] = "Do not infer missing source data.",
            ["htmlRule"] = "Render a structural review notice only.",
            ["openQuestions"] = new JsonArray()
        });
        var assignments = new JsonArray();
        foreach (var file in scanned.OrderBy(file => file.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            var groupId = file.RoutingState == "ASSIGNABLE" ? normalClusterOwner[file.LayoutClusterId].Id : exceptionOwner[file.LayoutClusterId].Id;
            assignments.Add(new JsonObject
            {
                ["relativePath"] = file.RelativePath,
                ["layoutClusterId"] = file.LayoutClusterId,
                ["semanticCategory"] = file.SemanticCategory,
                ["topLevelCategory"] = file.TopLevelCategory,
                ["routingState"] = file.RoutingState,
                ["groupId"] = groupId
            });
        }
        return new JsonObject
        {
            ["schemaVersion"] = PlanSchemaVersion,
            ["groups"] = jsonGroups,
            ["fileAssignments"] = assignments,
            ["unclassifiedFiles"] = new JsonArray(),
            ["provenance"] = new JsonObject
            {
                ["batchId"] = batchId,
                ["captureSource"] = "numeric-capture.sqlite",
                ["materializedUtc"] = DateTimeOffset.UtcNow.ToString("O"),
                ["scannedCount"] = scanned.Count,
                ["fallbackCount"] = fallbackCount,
                ["validationFile"] = "group-validation-report.json"
            }
        };
    }

    private static IReadOnlyList<SecondLevelGroup> BuildExceptionGroups(IReadOnlyList<TrackedWorkbook> scanned)
    {
        return scanned.Where(file => file.RoutingState is "EMPTY_LAYOUT" or "CAPTURE_INCOMPLETE")
            .GroupBy(file => file.LayoutClusterId, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal)
            .Select(group =>
            {
                var routes = group.Select(file => file.RoutingState).Distinct(StringComparer.Ordinal).OrderBy(route => route, StringComparer.Ordinal).ToList();
                var route = string.Join("_OR_", routes);
                var cluster = group.Key;
                var suffix = cluster.StartsWith("cluster-", StringComparison.Ordinal) ? cluster["cluster-".Length..] : cluster;
                var id = routes.SequenceEqual(["EMPTY_LAYOUT"], StringComparer.Ordinal) ? $"empty-layout-{suffix}" : $"capture-incomplete-{suffix}";
                var representatives = group.Select(file => file.RelativePath).Take(2).ToList();
                return new SecondLevelGroup(
                    id,
                    routes.SequenceEqual(["EMPTY_LAYOUT"], StringComparer.Ordinal) ? "Scanner-proven empty layout" : "Scanner-proven incomplete capture",
                    "renderNeedsReview",
                    "scanner-exception",
                    [cluster], routes, representatives,
                    ["Scanner routing state: " + route + ".", "No missing source data is inferred."],
                    "Do not infer tables or values that are absent from the captured source.",
                    "Render a named structural review notice with the captured source context.", []);
            }).ToList();
    }

    private static JsonObject GroupJson(SecondLevelGroup group, IReadOnlyDictionary<string, ClusterEvidence> evidenceByCluster) => new()
    {
        ["id"] = group.Id,
        ["name"] = group.Name,
        ["rendererKey"] = group.RendererKey,
        ["selector"] = new JsonObject { ["layoutClusterIds"] = new JsonArray(group.LayoutClusterIds.Select(id => (JsonNode?)id).ToArray()), ["fallback"] = false },
        ["componentRecipe"] = new JsonArray(group.ComponentRecipe.Select(item => (JsonNode?)item).ToArray()),
        ["representativeFiles"] = new JsonArray(NormalizedRepresentatives(group, evidenceByCluster).Select(path => (JsonNode?)path).ToArray()),
        ["importantData"] = new JsonArray(group.ImportantData.Select(item => (JsonNode?)item).ToArray()),
        ["extractionRule"] = group.ExtractionRule,
        ["htmlRule"] = group.HtmlRule,
        ["openQuestions"] = new JsonArray(group.OpenQuestions.Select(item => (JsonNode?)item).ToArray())
    };

    private static IReadOnlyList<string> NormalizedRepresentatives(SecondLevelGroup group, IReadOnlyDictionary<string, ClusterEvidence> evidenceByCluster)
    {
        var allowed = group.LayoutClusterIds.SelectMany(id => evidenceByCluster[id].RepresentativePaths).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var representatives = group.RepresentativeFiles.Where(path => allowed.Contains(path)).Distinct(StringComparer.OrdinalIgnoreCase).Take(2).ToList();
        if (representatives.Count == 0) representatives = group.LayoutClusterIds.SelectMany(id => evidenceByCluster[id].RepresentativePaths).Distinct(StringComparer.OrdinalIgnoreCase).Take(2).ToList();
        return representatives;
    }

    private static JsonObject BuildCatalog(JsonObject plan)
    {
        var catalog = (JsonObject)plan.DeepClone();
        foreach (var group in catalog["groups"]!.AsArray().Select(node => node!.AsObject()))
        {
            var selector = group["selector"]!.AsObject();
            var clusters = selector["layoutClusterIds"]?.AsArray().Select(node => node?.GetValue<string>()).Where(value => !string.IsNullOrWhiteSpace(value)).ToList() ?? [];
            group["memberSelectionRule"] = selector["fallback"]?.GetValue<bool>() == true
                ? "Fallback: EMPTY_LAYOUT or CAPTURE_INCOMPLETE only."
                : $"layout clusters = {string.Join("; ", clusters)}";
            group.Remove("selector");
        }
        return catalog;
    }

    private static async Task InvokeCodexAsync(string serviceDirectory, string outputPath, string prompt, Action<string>? log, CancellationToken cancellationToken)
    {
        var codex = AppRuntimePaths.Current.CodexExecutable;
        var schema = Path.Combine(serviceDirectory, "group-refinement.schema.json");
        if (!File.Exists(schema)) throw new FileNotFoundException("The compact group-refinement response schema is missing.", schema);
        var info = new ProcessStartInfo(codex)
        {
            WorkingDirectory = serviceDirectory,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardInputEncoding = Encoding.UTF8,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };
        info.ArgumentList.Add("exec"); info.ArgumentList.Add("--ephemeral"); info.ArgumentList.Add("--sandbox"); info.ArgumentList.Add("read-only");
        info.ArgumentList.Add("--model"); info.ArgumentList.Add(Environment.GetEnvironmentVariable("INFERENCE_DATA_AI_CODEX_MODEL") is { Length: > 0 } model ? model : "gpt-5.6-sol");
        info.ArgumentList.Add("--config"); info.ArgumentList.Add("model_reasoning_effort=xhigh");
        info.ArgumentList.Add("--output-schema"); info.ArgumentList.Add(schema); info.ArgumentList.Add("-o"); info.ArgumentList.Add(outputPath); info.ArgumentList.Add("-");
        var transcriptPath = outputPath + ".cli.log";
        var transcript = new StringBuilder();
        var transcriptLock = new object();
        void Record(string stream, string text)
        {
            lock (transcriptLock) transcript.Append('[').Append(DateTimeOffset.UtcNow.ToString("O")).Append("] [").Append(stream).Append("] ").AppendLine(text);
        }
        using var process = Process.Start(info) ?? throw new InvalidOperationException("Could not start Codex CLI for compact group refinement.");
        process.OutputDataReceived += (_, eventArgs) => { if (eventArgs.Data is not null) Record("OUT", eventArgs.Data); };
        process.ErrorDataReceived += (_, eventArgs) => { if (eventArgs.Data is not null) Record("ERR", eventArgs.Data); };
        process.BeginOutputReadLine(); process.BeginErrorReadLine();
        try
        {
            await process.StandardInput.WriteAsync(prompt);
            process.StandardInput.Close();
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromMinutes(10));
            await process.WaitForExitAsync(timeout.Token);
            if (process.ExitCode != 0) throw new InvalidOperationException($"Codex CLI compact group refinement exited with code {process.ExitCode}. See {transcriptPath}.");
            if (!File.Exists(outputPath)) throw new FileNotFoundException($"Codex CLI completed without writing compact group refinement output. See {transcriptPath}.", outputPath);
        }
        finally
        {
            string text;
            lock (transcriptLock) text = transcript.ToString();
            File.WriteAllText(transcriptPath, text, Utf8);
        }
    }

    private static string BuildPrompt(RefinementWorkItem workItem, IReadOnlyList<SecondLevelGroup> knownGroups, IReadOnlyList<string> retryErrors)
    {
        var evidence = new
        {
            schemaVersion = "compact-cluster-evidence-v1",
            topLevelCategory = workItem.TopLevelCategory,
            compactBatch = workItem.Ordinal,
            clusters = workItem.Clusters.Select(cluster => new
            {
                layoutClusterId = cluster.LayoutClusterId,
                fileCount = cluster.FileCount,
                sectionCandidateSequence = cluster.SectionRecipe.Select(section => section.CandidateType),
                normalizedTitleFacets = cluster.SectionRecipe.Select(section => section.Title),
                normalizedHeaderFacets = cluster.SectionRecipe.Select(section => section.Headers),
                normalizedLogicalRowFacets = cluster.SectionRecipe.Select(section => section.LogicalRows),
                tableCount = cluster.TableCount,
                tableShape = cluster.TableShape,
                representativePaths = cluster.RepresentativePaths
            }),
            knownSecondLevelGroups = knownGroups.Select(group => new { id = group.Id, name = group.Name, rendererKey = group.RendererKey, componentRecipe = group.ComponentRecipe }),
            retryValidationErrors = retryErrors
        };
        var expectedRenderer = RendererForTopLevel[workItem.TopLevelCategory];
        return "You are the second-level Excel form grouping stage. The top-level category below is only a partition, never a final renderer group.\n"
            + "Return JSON matching the supplied schema and assign every listed layoutClusterId exactly once to a genuinely similar named second-level group.\n"
            + "Do not use a fallback group. The pipeline reserves fallback exclusively for scanner-proven EMPTY_LAYOUT or CAPTURE_INCOMPLETE files and does not send those clusters here.\n"
            + "Use only the required renderer key for this top-level category: " + expectedRenderer + ".\n"
            + "For every group, componentRecipe must exactly equal the ordered distinct sectionCandidateSequence of every cluster assigned to it. Split any clusters whose sequence differs. Normalize group ids as stable lowercase slugs. Reuse an earlier group id only when it has the same renderer and componentRecipe; otherwise create a new descriptive id. Representative paths must be copied from the supplied evidence, never invented. Do not inspect local files, use shell/web/GitHub/tools, or ask for more inputs.\n\n"
            // Codex writes the supplied prompt to stderr. Keep the payload compact
            // so the pipe reader is not needlessly flooded with thousands of
            // formatted JSON lines; the complete transport transcript is still
            // retained beside the response for diagnosis.
            + "Compact evidence JSON:\n" + JsonSerializer.Serialize(evidence);
    }

    private static RefinementResponse ParseRefinementResponse(string path, string expectedTopLevel, IReadOnlyList<ClusterEvidence> expectedClusters)
    {
        var root = JsonNode.Parse(File.ReadAllText(path, Encoding.UTF8))?.AsObject() ?? throw new InvalidOperationException("Compact group refinement response is not a JSON object.");
        if (!string.Equals(root["schemaVersion"]?.GetValue<string>(), RefinementSchemaVersion, StringComparison.Ordinal))
            throw new InvalidOperationException($"Compact group refinement response must use schemaVersion '{RefinementSchemaVersion}'.");
        var category = root["topLevelCategory"]?.GetValue<string>() ?? string.Empty;
        var groups = root["groups"]?.AsArray() ?? throw new InvalidOperationException("Compact group refinement response has no groups.");
        var values = new List<SecondLevelGroup>();
        foreach (var node in groups)
        {
            var group = node?.AsObject() ?? throw new InvalidOperationException("Compact group refinement contains an invalid group.");
            string String(string key) => group[key]?.GetValue<string>() ?? string.Empty;
            List<string> Strings(string key) => group[key]?.AsArray().Select(value => value?.GetValue<string>() ?? string.Empty).ToList() ?? [];
            values.Add(new SecondLevelGroup(
                String("id"), String("name"), String("rendererKey"), expectedTopLevel,
                Strings("layoutClusterIds"), Strings("componentRecipe"), Strings("representativeFiles"), Strings("importantData"),
                String("extractionRule"), String("htmlRule"), Strings("openQuestions")));
        }
        return new RefinementResponse(category, values);
    }

    private static IReadOnlyList<RefinementWorkItem> BuildWorkItems(IReadOnlySet<string> assignableClusters, IReadOnlyDictionary<string, SemanticCategory> categories, IReadOnlyDictionary<string, ClusterEvidence> evidence, int batchSize)
    {
        var items = new List<RefinementWorkItem>();
        foreach (var topLevel in assignableClusters.GroupBy(cluster => categories[cluster].TopLevelCategory, StringComparer.Ordinal).OrderBy(group => group.Key, StringComparer.Ordinal))
        {
            var clusters = topLevel.Select(cluster => evidence[cluster]).OrderByDescending(cluster => cluster.FileCount).ThenBy(cluster => cluster.LayoutClusterId, StringComparer.Ordinal).ToList();
            for (var index = 0; index < clusters.Count; index += batchSize)
                items.Add(new RefinementWorkItem(topLevel.Key, index / batchSize + 1, clusters.Skip(index).Take(batchSize).ToList()));
        }
        return items;
    }

    private static Dictionary<string, ClusterEvidence> ReadClusterEvidence(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        var values = new Dictionary<string, ClusterEvidence>(StringComparer.Ordinal);
        foreach (var cluster in document.RootElement.GetProperty("clusters").EnumerateArray())
        {
            var id = cluster.GetProperty("LayoutClusterId").GetString() ?? string.Empty;
            var recipe = cluster.GetProperty("sectionRecipe").EnumerateArray().Select(section => new SectionEvidence(
                section.GetProperty("CandidateType").GetString() ?? string.Empty,
                ReadStringArray(section, "title"), ReadStringArray(section, "headers"), ReadStringArray(section, "logicalRowFacets"))).ToList();
            var shape = cluster.TryGetProperty("tableShape", out var shapeElement)
                ? shapeElement.EnumerateArray().Select(item => new TableShapeEvidence(item.GetProperty("candidateType").GetString() ?? string.Empty, item.GetProperty("rowSpan").GetString() ?? string.Empty, item.GetProperty("columnSpan").GetString() ?? string.Empty)).ToList()
                : [];
            values.Add(id, new ClusterEvidence(id, cluster.GetProperty("FileCount").GetInt32(),
                cluster.TryGetProperty("representativePaths", out var paths) ? paths.EnumerateArray().Select(item => item.GetString() ?? string.Empty).Where(path => path.Length > 0).Take(2).ToList() : [],
                cluster.TryGetProperty("tableCount", out var count) ? count.GetInt32() : recipe.Count,
                shape, recipe));
        }
        return values;
    }

    private static Dictionary<string, SemanticCategory> ReadSemanticCategories(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        var values = new Dictionary<string, SemanticCategory>(StringComparer.Ordinal);
        foreach (var category in document.RootElement.GetProperty("categories").EnumerateArray())
        {
            var name = category.GetProperty("category").GetString() ?? string.Empty;
            var topLevel = category.TryGetProperty("topLevelCategory", out var topLevelElement) ? topLevelElement.GetString() ?? string.Empty : TopLevelCategory(name);
            foreach (var cluster in category.GetProperty("layoutClusterIds").EnumerateArray()) values.Add(cluster.GetString() ?? string.Empty, new SemanticCategory(name, topLevel));
        }
        return values;
    }

    private static Dictionary<string, string> ReadLayoutClusters(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        return document.RootElement.GetProperty("clusters").EnumerateArray().SelectMany(cluster =>
            cluster.GetProperty("MemberPaths").EnumerateArray().Select(member => (Path: member.GetString() ?? string.Empty, Cluster: cluster.GetProperty("LayoutClusterId").GetString() ?? string.Empty)))
            .ToDictionary(item => item.Path, item => item.Cluster, StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, InventoryDocument> ReadInventory(string path)
    {
        var values = new Dictionary<string, InventoryDocument>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in File.ReadLines(path, Encoding.UTF8))
        {
            using var document = JsonDocument.Parse(line);
            var root = document.RootElement;
            values.Add(root.GetProperty("RelativePath").GetString() ?? string.Empty, new InventoryDocument(root.GetProperty("RoutingState").GetString() ?? string.Empty));
        }
        return values;
    }

    private static List<string> ReadClassificationPaths(string path)
    {
        using var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        var header = ParseCsvRow(reader.ReadLine() ?? string.Empty);
        var columns = header.Select((name, index) => (name, index)).ToDictionary(item => item.name, item => item.index, StringComparer.OrdinalIgnoreCase);
        var values = new List<string>();
        while (reader.ReadLine() is { } line)
        {
            var row = ParseCsvRow(line);
            string Value(string name) => columns.TryGetValue(name, out var index) && index < row.Count ? row[index] : string.Empty;
            if (string.Equals(Value("status"), "SCANNED", StringComparison.OrdinalIgnoreCase)) values.Add(Value("relativePath"));
        }
        if (values.Any(string.IsNullOrWhiteSpace) || values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != values.Count)
            throw new InvalidOperationException("Structure classification has missing or duplicate scanned relative paths.");
        return values;
    }

    private static int ReadScannedCount(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
        return document.RootElement.GetProperty("statusCounts").GetProperty("SCANNED").GetInt32();
    }

    private static List<string> ParseCsvRow(string line)
    {
        var values = new List<string>(); var value = new StringBuilder(); var quoted = false;
        for (var index = 0; index < line.Length; index++)
        {
            var character = line[index];
            if (character == '"' && quoted && index + 1 < line.Length && line[index + 1] == '"') { value.Append(character); index++; continue; }
            if (character == '"') { quoted = !quoted; continue; }
            if (character == ',' && !quoted) { values.Add(value.ToString()); value.Clear(); continue; }
            value.Append(character);
        }
        values.Add(value.ToString()); return values;
    }

    private static IReadOnlyList<string> ReadStringArray(JsonElement element, string property) => element.TryGetProperty(property, out var array)
        ? array.EnumerateArray().Select(value => value.GetString() ?? string.Empty).Where(value => value.Length > 0).ToList() : [];
    private static string TopLevelCategory(string category) { var separator = category.IndexOf("--", StringComparison.Ordinal); return separator < 0 ? category : category[..separator]; }
    private static string ComponentRecipeFor(ClusterEvidence cluster) => string.Join("|", cluster.SectionRecipe.Select(section => section.CandidateType).Distinct(StringComparer.Ordinal));
    private static bool SequenceEqual(IReadOnlyList<string> first, IReadOnlyList<string> second) => first.SequenceEqual(second, StringComparer.Ordinal);
    private static int ClustersPerRequest() => int.TryParse(Environment.GetEnvironmentVariable("INFERENCE_DATA_AI_GROUP_BATCH_SIZE"), out var value) && value is >= 8 and <= 60 ? value : DefaultClustersPerRequest;
    private static string SafeFilePart(string value) => Regex.Replace(value, "[^a-zA-Z0-9._-]+", "-");
    private static void AtomicWrite(string path, string text) { var temporary = path + ".tmp"; File.WriteAllText(temporary, text, Utf8); File.Move(temporary, path, overwrite: true); }
    private static readonly JsonSerializerOptions Indented = new() { WriteIndented = true };

    private sealed record InventoryDocument(string RoutingState);
    private sealed record SemanticCategory(string Category, string TopLevelCategory);
    private sealed record TrackedWorkbook(string RelativePath, string LayoutClusterId, string SemanticCategory, string TopLevelCategory, string RoutingState);
    private sealed record SectionEvidence(string CandidateType, IReadOnlyList<string> Title, IReadOnlyList<string> Headers, IReadOnlyList<string> LogicalRows);
    private sealed record TableShapeEvidence(string CandidateType, string RowSpan, string ColumnSpan);
    private sealed record ClusterEvidence(string LayoutClusterId, int FileCount, IReadOnlyList<string> RepresentativePaths, int TableCount, IReadOnlyList<TableShapeEvidence> TableShape, IReadOnlyList<SectionEvidence> SectionRecipe);
    private sealed record RefinementWorkItem(string TopLevelCategory, int Ordinal, IReadOnlyList<ClusterEvidence> Clusters);
    private sealed record RefinementResponse(string TopLevelCategory, IReadOnlyList<SecondLevelGroup> Groups);
    private sealed record RefinementBatchResult(IReadOnlyList<SecondLevelGroup> Groups, string ValidationPath);
    private sealed record MergedValidation(IReadOnlyList<string> Errors, int FallbackCount);
    private sealed class SecondLevelGroup
    {
        internal SecondLevelGroup(string id, string name, string rendererKey, string topLevelCategory, IReadOnlyList<string> layoutClusterIds, IReadOnlyList<string> componentRecipe, IReadOnlyList<string> representativeFiles, IReadOnlyList<string> importantData, string extractionRule, string htmlRule, IReadOnlyList<string> openQuestions)
        {
            Id = id; Name = name; RendererKey = rendererKey; TopLevelCategory = topLevelCategory;
            // Keep duplicate ids until the validation boundary can report them to
            // the bounded retry instead of silently normalizing an invalid answer.
            LayoutClusterIds = layoutClusterIds.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            // The response contract uses an ordered *distinct* candidate recipe.
            // Some otherwise-valid model outputs repeat a candidate for repeated
            // sections, so normalize only that redundant notation while preserving
            // order; validation below still rejects genuinely different recipes.
            ComponentRecipe = componentRecipe.Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.Ordinal).ToList();
            RepresentativeFiles = representativeFiles.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            ImportantData = importantData.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            ExtractionRule = extractionRule; HtmlRule = htmlRule; OpenQuestions = openQuestions.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
        }
        internal string Id { get; }
        internal string Name { get; }
        internal string RendererKey { get; }
        internal string TopLevelCategory { get; }
        internal List<string> LayoutClusterIds { get; }
        internal List<string> ComponentRecipe { get; }
        internal List<string> RepresentativeFiles { get; }
        internal List<string> ImportantData { get; }
        internal string ExtractionRule { get; }
        internal string HtmlRule { get; }
        internal List<string> OpenQuestions { get; }
        internal void Merge(SecondLevelGroup other)
        {
            LayoutClusterIds.AddRange(other.LayoutClusterIds.Where(value => !LayoutClusterIds.Contains(value, StringComparer.Ordinal)));
            RepresentativeFiles.AddRange(other.RepresentativeFiles.Where(value => !RepresentativeFiles.Contains(value, StringComparer.OrdinalIgnoreCase)));
        }
    }
}

internal sealed record TwoLevelGroupAnalysisRequest(string ServiceDirectory, string BatchId);
internal sealed record TwoLevelGroupAnalysisRunResult(string BatchDirectory, string PlanPath, string CatalogPath, string ValidationPath, int GroupCount, int ScannedCount, int FallbackCount);
