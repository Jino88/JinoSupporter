using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;

namespace InferenceDataAIService.Wpf;

internal sealed class CanonicalEvidenceClient(AppPathSettings paths)
{
    private readonly AppPathSettings _paths = paths;
    private readonly string _serviceDirectory =
        Path.GetFullPath(paths.ServiceDirectory);

    internal async Task<EvidenceAnswerSession> AskAsync(
        string databasePath,
        string question,
        string queryMode = "canonical")
    {
        var outputDirectory = _paths.EvidenceOutputDirectory;
        Directory.CreateDirectory(outputDirectory);
        var token = $"{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}_{Guid.NewGuid():N}";
        var answerPath = Path.Combine(outputDirectory, $"{token}.answer.json");
        var markdownPath = Path.Combine(outputDirectory, $"{token}.answer.md");
        var historyDatabasePath = HistoryDatabasePath();
        if (string.Equals(
                queryMode,
                "history",
                StringComparison.OrdinalIgnoreCase))
        {
            if (!File.Exists(historyDatabasePath))
                throw new FileNotFoundException(
                    "기존 전체 이력 DB를 찾을 수 없습니다.",
                    historyDatabasePath);
            var requestPath = Path.Combine(
                outputDirectory,
                $"{token}.relevance-request.json");
            await RunCliAsync(
                "table-first-relevance-query",
                "--db", historyDatabasePath,
                "--question", question,
                "--candidate-limit", "200",
                "--reasoning-effort", "medium",
                "--out-request", requestPath,
                "--out-json", answerPath,
                "--out-markdown", markdownPath);
        }
        else
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException(
                    "최신 적재 canonical DB를 찾을 수 없습니다.",
                    databasePath);
            await RunCliAsync(
                "evidence-answer",
                "--db", Path.GetFullPath(databasePath),
                "--question", question,
                "--out-json", answerPath,
                "--out-markdown", markdownPath);
        }
        return EvidenceAnswerSession.Load(answerPath, markdownPath);
    }

    internal async Task<EvidenceDetailDocument> DetailAsync(
        string databasePath,
        string evidenceId)
    {
        var outputDirectory = _paths.EvidenceDetailDirectory;
        Directory.CreateDirectory(outputDirectory);
        var safeId = string.Concat(
            evidenceId.Select(character =>
                char.IsLetterOrDigit(character) || character == '-'
                    ? character
                    : '_'));
        var outputPath = Path.Combine(
            outputDirectory,
            $"{safeId}_{Guid.NewGuid():N}.json");
        if (evidenceId.StartsWith("TF-EVD-", StringComparison.OrdinalIgnoreCase))
        {
            var historyDatabasePath = HistoryDatabasePath();
            if (!File.Exists(historyDatabasePath))
                throw new FileNotFoundException(
                    "전체 이력 검색 DB를 찾을 수 없습니다.",
                    historyDatabasePath);
            await RunCliAsync(
                "table-first-history-detail",
                "--db", historyDatabasePath,
                "--evidence-id", evidenceId,
                "--out", outputPath);
        }
        else
        {
            await RunCliAsync(
                "evidence-detail",
                "--db", Path.GetFullPath(databasePath),
                "--evidence-id", evidenceId,
                "--out", outputPath);
        }
        return EvidenceDetailDocument.Load(outputPath);
    }

    internal string HistoryDatabasePath() =>
        _paths.HistoryDatabasePath;

    internal EvidenceAnswerSession? LoadLatestRelevanceSession()
    {
        var directories = new[]
        {
            _paths.EvidenceOutputDirectory,
            _paths.TableFirstRelevanceDirectory,
        };
        var candidates = new List<(string JsonPath, string MarkdownPath, int RawValueCount, DateTime LastWriteTime)>();
        foreach (var directory in directories)
        {
            if (!Directory.Exists(directory)) continue;
            foreach (var jsonPath in Directory.EnumerateFiles(
                         directory,
                         "*.answer.json",
                         SearchOption.TopDirectoryOnly))
            {
                try
                {
                    using var document = JsonDocument.Parse(
                        File.ReadAllText(jsonPath, Encoding.UTF8));
                    var root = document.RootElement;
                    if (!root.TryGetProperty("schemaVersion", out var schema)
                        || !string.Equals(
                            schema.GetString(),
                            "table-first-relevance-result-v1",
                            StringComparison.Ordinal))
                        continue;
                    var markdownPath = Path.ChangeExtension(jsonPath, ".md");
                    if (!File.Exists(markdownPath)) continue;
                    var rawValueCount = root.TryGetProperty(
                            "coverage",
                            out var coverage)
                        && coverage.TryGetProperty(
                            "rawDataPointCount",
                            out var rawDataPointCount)
                        && rawDataPointCount.TryGetInt32(out var count)
                            ? count
                            : 0;
                    candidates.Add((
                        jsonPath,
                        markdownPath,
                        rawValueCount,
                        File.GetLastWriteTimeUtc(jsonPath)));
                }
                catch (JsonException)
                {
                    // Ignore partial or incompatible historical artifacts.
                }
            }
        }
        var selected = candidates
            .OrderByDescending(item => item.RawValueCount > 0)
            .ThenByDescending(item => item.LastWriteTime)
            .FirstOrDefault();
        return string.IsNullOrWhiteSpace(selected.JsonPath)
            ? null
            : EvidenceAnswerSession.Load(
                selected.JsonPath,
                selected.MarkdownPath);
    }

    internal async Task<IngestWorkbookResult> IngestAsync(
        string databasePath,
        string workbookPath,
        bool retryFailed,
        IProgress<IngestProgressEvent>? progress = null,
        bool inspectAuthDialog = false,
        bool dismissAuthDialog = false,
        string authDialogTitle = "",
        string authDialogClass = "",
        string authDialogButton = "")
    {
        var artifactRoot = _paths.IncrementalIngestDirectory;
        var arguments = new List<string>
        {
            "ingest-workbook",
            "--db", Path.GetFullPath(databasePath),
            "--input", Path.GetFullPath(workbookPath),
            "--artifact-root", artifactRoot,
            "--dataset", AppPathSettings.DefaultDataset,
            "--capture-backend", "com",
            "--covered-cell-mode", "blank",
        };
        if (retryFailed)
        {
            arguments.Add("--repair-rejected-draft");
            arguments.Add("--repair-unselected-source");
        }
        AddAuthDialogArguments(
            arguments,
            inspectAuthDialog,
            dismissAuthDialog,
            authDialogTitle,
            authDialogClass,
            authDialogButton);
        var result = await RunCliAsync(progress, arguments.ToArray());
        using var document = JsonDocument.Parse(result.StandardOutput);
        var root = document.RootElement;
        return new IngestWorkbookResult(
            GetString(root, "status"),
            GetString(root, "workbookStatus"),
            GetString(root, "revisionUid"),
            GetString(root, "sourcePath"),
            GetString(root, "publicAnalysisId"),
            GetString(root, "journalPath"),
            GetString(root, "manifestPath"),
            GetString(root, "artifactDirectory"),
            root.TryGetProperty("studies", out var studies)
                && studies.TryGetInt32(out var studyCount)
                    ? studyCount
                    : 0,
            root.TryGetProperty("integrityOk", out var integrityOk)
                && integrityOk.ValueKind is JsonValueKind.True
                    or JsonValueKind.False
                    ? integrityOk.GetBoolean()
                    : null,
            result.StandardOutput);
    }

    internal async Task<IngestCorpusResult> IngestCorpusAsync(
        string databasePath,
        string sourceDirectory,
        bool retryFailed,
        IProgress<IngestProgressEvent>? progress = null,
        bool inspectAuthDialog = false,
        bool dismissAuthDialog = false,
        string authDialogTitle = "",
        string authDialogClass = "",
        string authDialogButton = "",
        string sourceManifestPath = "")
    {
        var artifactRoot = _paths.CorpusIngestDirectory;
        var arguments = new List<string>
        {
            "ingest-corpus",
            "--db", Path.GetFullPath(databasePath),
            "--input", Path.GetFullPath(sourceDirectory),
            "--artifact-root", artifactRoot,
            "--dataset", AppPathSettings.DefaultDataset,
            "--capture-backend", "com",
            "--covered-cell-mode", "blank",
            "--workbook-workers", "4",
            "--com-workers", "1",
            "--packet-workers", "3",
            "--ai-workers", "3",
            "--db-workers", "1",
        };
        if (!string.IsNullOrWhiteSpace(sourceManifestPath))
        {
            arguments.Add("--source-manifest");
            arguments.Add(Path.GetFullPath(sourceManifestPath));
        }
        if (retryFailed)
        {
            arguments.Add("--retry-failed");
            arguments.Add("--repair-rejected-draft");
            arguments.Add("--repair-unselected-source");
        }
        AddAuthDialogArguments(
            arguments,
            inspectAuthDialog,
            dismissAuthDialog,
            authDialogTitle,
            authDialogClass,
            authDialogButton);
        var result = await RunCliAsync(progress, arguments.ToArray());
        using var document = JsonDocument.Parse(result.StandardOutput);
        var root = document.RootElement;
        return new IngestCorpusResult(
            GetString(root, "status"),
            GetString(root, "sourceRoot"),
            GetString(root, "journal"),
            GetString(root, "result"),
            GetInt32(root, "selectedSources"),
            GetInt32(root, "completedThisRun"),
            GetInt32(root, "failedThisRun"),
            result.StandardOutput);
    }

    internal async Task<FormPipelineCompleteResult>
        CompleteFormPipelineAsync(
            string databasePath,
            string sourceDirectory,
            string reviewer,
            IProgress<IngestProgressEvent>? progress = null)
    {
        if (string.IsNullOrWhiteSpace(reviewer))
            throw new ArgumentException(
                "Reviewer is required.",
                nameof(reviewer));
        var result = await RunCliAsync(
            progress,
            "form-pipeline-complete",
            "--db", Path.GetFullPath(databasePath),
            "--input", Path.GetFullPath(sourceDirectory),
            "--output-root",
            Path.GetFullPath(_paths.OutputRootDirectory),
            "--reviewer", reviewer.Trim(),
            "--dataset", AppPathSettings.DefaultDataset,
            "--analysis-workers", "4",
            "--reasoning-effort", "medium",
            "--analysis-timeout", "900",
            "--com-timeout-seconds", "120");
        using var document = JsonDocument.Parse(
            result.StandardOutput);
        var root = document.RootElement;
        var preflight = root.GetProperty("preflight");
        var formGroups = root.GetProperty("formGroups");
        var corpus = root.TryGetProperty(
                "corpus",
                out var corpusElement)
            && corpusElement.ValueKind == JsonValueKind.Object
                ? corpusElement
                : default;
        return new FormPipelineCompleteResult(
            GetString(root, "status"),
            GetString(root, "result"),
            GetString(root, "report"),
            GetString(root, "review"),
            GetString(root, "manifest"),
            GetInt32(preflight, "total"),
            GetInt32(preflight, "knownForms"),
            GetInt32(preflight, "excludedForms"),
            GetInt32(preflight, "captureFailed"),
            GetInt32(formGroups, "groupCount"),
            GetInt32(formGroups, "pendingCount"),
            GetInt32(formGroups, "approvedCount"),
            GetInt32(root, "analysisErrors"),
            corpus.ValueKind == JsonValueKind.Object
                ? GetInt32(corpus, "selectedSources")
                : 0,
            corpus.ValueKind == JsonValueKind.Object
                ? GetInt32(corpus, "attempted")
                : 0,
            corpus.ValueKind == JsonValueKind.Object
                ? GetInt32(corpus, "completedThisRun")
                : 0,
            corpus.ValueKind == JsonValueKind.Object
                ? GetInt32(corpus, "failedThisRun")
                : 0,
            result.StandardOutput);
    }

    internal async Task<FormPreflightDocument> PreflightFormsAsync(
        string databasePath,
        string sourceDirectory,
        IProgress<IngestProgressEvent>? progress = null,
        bool inspectAuthDialog = false,
        bool dismissAuthDialog = false,
        string authDialogTitle = "",
        string authDialogClass = "",
        string authDialogButton = "")
    {
        var outputDirectory = _paths.FormPreflightDirectory;
        Directory.CreateDirectory(outputDirectory);
        var outputPath = Path.Combine(
            outputDirectory,
            "latest.json");
        var cancelPath = FormPreflightCancelPath();
        File.Delete(cancelPath);
        var arguments = new List<string>
        {
            "form-preflight",
            "--db", Path.GetFullPath(databasePath),
            "--input", Path.GetFullPath(sourceDirectory),
            "--out", outputPath,
            "--output-root",
            Path.GetFullPath(_paths.OutputRootDirectory),
            "--cancel-file", cancelPath,
            "--dataset", AppPathSettings.DefaultDataset,
        };
        AddAuthDialogArguments(
            arguments,
            inspectAuthDialog,
            dismissAuthDialog,
            authDialogTitle,
            authDialogClass,
            authDialogButton);
        try
        {
            await RunCliAsync(progress, arguments.ToArray());
            return FormPreflightDocument.Load(outputPath);
        }
        finally
        {
            try
            {
                File.Delete(cancelPath);
            }
            catch (IOException)
            {
                // The next run also removes a stale marker before starting.
            }
            catch (UnauthorizedAccessException)
            {
                // The next run also removes a stale marker before starting.
            }
        }
    }

    internal void RequestFormPreflightCancellation()
    {
        var cancelPath = FormPreflightCancelPath();
        Directory.CreateDirectory(
            Path.GetDirectoryName(cancelPath)
            ?? _paths.FormPreflightDirectory);
        File.WriteAllText(
            cancelPath,
            DateTime.UtcNow.ToString("O"),
            Encoding.UTF8);
    }

    private string FormPreflightCancelPath() =>
        Path.Combine(
            _paths.FormPreflightDirectory,
            "cancel.request");

    internal FormPreflightDocument? LoadLatestFormPreflight()
    {
        var path = Path.Combine(
            _paths.FormPreflightDirectory,
            "latest.json");
        try
        {
            return File.Exists(path)
                ? FormPreflightDocument.Load(path)
                : null;
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or JsonException
            or InvalidDataException)
        {
            return null;
        }
    }

    private string FormGroupReviewPath() =>
        Path.Combine(
            _paths.FormPreflightDirectory,
            "group-review.latest.json");

    internal FormGroupReviewDocument? LoadLatestFormGroupReview()
    {
        var path = FormGroupReviewPath();
        try
        {
            return File.Exists(path)
                ? FormGroupReviewDocument.Load(path)
                : null;
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or JsonException
            or InvalidDataException)
        {
            return null;
        }
    }

    internal async Task<FormGroupReviewDocument> RefreshFormGroupsAsync(
        string databasePath,
        string preflightReportPath)
    {
        var reviewPath = FormGroupReviewPath();
        await RunCliAsync(
            "form-group-review",
            "--db", Path.GetFullPath(databasePath),
            "--report", Path.GetFullPath(preflightReportPath),
            "--out", reviewPath,
            "--output-root",
            Path.GetFullPath(_paths.OutputRootDirectory));
        return FormGroupReviewDocument.Load(reviewPath);
    }

    internal async Task<FormGroupReviewDocument> AnalyzeFormFamilyAsync(
        string databasePath,
        string preflightReportPath,
        string familyId)
    {
        if (string.IsNullOrWhiteSpace(familyId))
            throw new ArgumentException(
                "Family ID is required.",
                nameof(familyId));
        var reviewPath = FormGroupReviewPath();
        await RunCliAsync(
            "form-family-analyze",
            "--db", Path.GetFullPath(databasePath),
            "--report", Path.GetFullPath(preflightReportPath),
            "--family-id", familyId.Trim(),
            "--review-out", reviewPath,
            "--output-root",
            Path.GetFullPath(_paths.OutputRootDirectory),
            "--reasoning-effort", "medium");
        return FormGroupReviewDocument.Load(reviewPath);
    }

    internal async Task<FormGroupReviewDocument> DecideFormFamilyAsync(
        string databasePath,
        string preflightReportPath,
        string familyId,
        string decision,
        string reviewer,
        string displayName = "",
        string linkedFormSignatureId = "",
        string notes = "")
    {
        if (string.IsNullOrWhiteSpace(familyId))
            throw new ArgumentException(
                "Family ID is required.",
                nameof(familyId));
        if (string.IsNullOrWhiteSpace(reviewer))
            throw new ArgumentException(
                "Reviewer is required.",
                nameof(reviewer));
        var reviewPath = FormGroupReviewPath();
        var arguments = new List<string>
        {
            "form-family-decide",
            "--db", Path.GetFullPath(databasePath),
            "--report", Path.GetFullPath(preflightReportPath),
            "--family-id", familyId.Trim(),
            "--decision", decision.Trim().ToUpperInvariant(),
            "--reviewer", reviewer.Trim(),
            "--review-out", reviewPath,
            "--output-root",
            Path.GetFullPath(_paths.OutputRootDirectory),
        };
        if (!string.IsNullOrWhiteSpace(displayName))
            arguments.AddRange(
            [
                "--display-name", displayName.Trim(),
            ]);
        if (!string.IsNullOrWhiteSpace(linkedFormSignatureId))
            arguments.AddRange(
            [
                "--linked-form-signature-id",
                linkedFormSignatureId.Trim(),
            ]);
        if (!string.IsNullOrWhiteSpace(notes))
            arguments.AddRange(["--notes", notes.Trim()]);
        await RunCliAsync(arguments.ToArray());
        return FormGroupReviewDocument.Load(reviewPath);
    }

    internal IngestWorkbookResult? LoadLatestIngestResult(
        string workbookPath)
    {
        if (string.IsNullOrWhiteSpace(workbookPath))
            return null;
        string sourcePath;
        try
        {
            sourcePath = Path.GetFullPath(workbookPath);
        }
        catch (Exception exception) when (
            exception is ArgumentException
            or NotSupportedException
            or PathTooLongException)
        {
            return null;
        }

        var artifactRoot = _paths.IncrementalIngestDirectory;
        if (!Directory.Exists(artifactRoot)) return null;
        IEnumerable<string> journalPaths;
        try
        {
            journalPaths = Directory.EnumerateFiles(
                    artifactRoot,
                    "journal.json",
                    SearchOption.AllDirectories)
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .ToList();
        }
        catch (Exception exception) when (
            exception is IOException
            or UnauthorizedAccessException
            or DirectoryNotFoundException)
        {
            return null;
        }

        foreach (var journalPath in journalPaths)
        {
            try
            {
                var rawJournal = File.ReadAllText(
                    journalPath,
                    Encoding.UTF8);
                using var document = JsonDocument.Parse(rawJournal);
                var root = document.RootElement;
                if (!root.TryGetProperty("result", out var result)
                    || result.ValueKind != JsonValueKind.Object)
                    continue;
                var resultSourcePath = GetString(result, "sourcePath");
                if (string.IsNullOrWhiteSpace(resultSourcePath)
                    && root.TryGetProperty("source", out var source)
                    && source.ValueKind == JsonValueKind.Object)
                    resultSourcePath = GetString(source, "sourcePath");
                if (string.IsNullOrWhiteSpace(resultSourcePath)
                    || !string.Equals(
                        Path.GetFullPath(resultSourcePath),
                        sourcePath,
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                var status = GetString(result, "status");
                if (string.IsNullOrWhiteSpace(status))
                    status = GetString(root, "status");
                if (string.Equals(
                        status,
                        "FAILED",
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        status,
                        "RUNNING",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                var artifactDirectory = GetString(
                    result,
                    "artifactDirectory");
                var manifestPath = GetString(result, "manifestPath");
                if (string.IsNullOrWhiteSpace(manifestPath)
                    && !string.IsNullOrWhiteSpace(artifactDirectory))
                    manifestPath = Path.Combine(
                        artifactDirectory,
                        "canonical-study-manifest.json");
                bool? integrityOk = null;
                if (result.TryGetProperty(
                        "integrityOk",
                        out var integrity)
                    && integrity.ValueKind is JsonValueKind.True
                        or JsonValueKind.False)
                    integrityOk = integrity.GetBoolean();
                return new IngestWorkbookResult(
                    status,
                    GetString(result, "workbookStatus"),
                    GetString(result, "revisionUid"),
                    resultSourcePath,
                    GetString(result, "publicAnalysisId"),
                    journalPath,
                    manifestPath,
                    artifactDirectory,
                    GetInt32(result, "studies"),
                    integrityOk,
                    result.GetRawText());
            }
            catch (Exception exception) when (
                exception is IOException
                or UnauthorizedAccessException
                or JsonException
                or ArgumentException
                or NotSupportedException)
            {
                // Ignore unrelated partial or inaccessible journals.
            }
        }
        return null;
    }

    private static void AddAuthDialogArguments(
        List<string> arguments,
        bool inspectAuthDialog,
        bool dismissAuthDialog,
        string authDialogTitle,
        string authDialogClass,
        string authDialogButton)
    {
        if (inspectAuthDialog) arguments.Add("--inspect-auth-dialog");
        if (!dismissAuthDialog) return;
        arguments.Add("--dismiss-auth-dialog");
        arguments.AddRange(
        [
            "--auth-dialog-title", authDialogTitle,
            "--auth-dialog-class", authDialogClass,
            "--auth-dialog-button", authDialogButton,
        ]);
    }

    internal async Task<RelatedStudiesDocument> RelatedAsync(
        string databasePath,
        string targetIdentifier,
        int limit = 12)
    {
        var outputDirectory = _paths.RelatedStudiesDirectory;
        Directory.CreateDirectory(outputDirectory);
        var outputPath = Path.Combine(
            outputDirectory,
            $"{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}_{Guid.NewGuid():N}.json");
        await RunCliAsync(
            "related-studies",
            "--db", Path.GetFullPath(databasePath),
            "--target", targetIdentifier,
            "--limit", limit.ToString(
                System.Globalization.CultureInfo.InvariantCulture),
            "--out", outputPath);
        return RelatedStudiesDocument.Load(outputPath);
    }

    internal async Task<ReviewQueueDocument> ReviewQueueAsync(
        string databasePath,
        int limit = 500)
    {
        var outputDirectory = _paths.HumanReviewDirectory;
        Directory.CreateDirectory(outputDirectory);
        var outputPath = Path.Combine(
            outputDirectory,
            $"queue_{DateTime.UtcNow:yyyyMMddTHHmmssfffZ}_{Guid.NewGuid():N}.json");
        await RunCliAsync(
            "review-queue",
            "--db", Path.GetFullPath(databasePath),
            "--limit", limit.ToString(
                System.Globalization.CultureInfo.InvariantCulture),
            "--out", outputPath);
        return ReviewQueueDocument.Load(outputPath);
    }

    internal async Task<ReviewDetailDocument> ReviewDetailAsync(
        string databasePath,
        string comparisonId)
    {
        var outputDirectory = _paths.HumanReviewDirectory;
        Directory.CreateDirectory(outputDirectory);
        var outputPath = Path.Combine(
            outputDirectory,
            $"{comparisonId}_{Guid.NewGuid():N}.json");
        await RunCliAsync(
            "review-detail",
            "--db", Path.GetFullPath(databasePath),
            "--comparison-id", comparisonId,
            "--out", outputPath);
        return ReviewDetailDocument.Load(outputPath);
    }

    internal async Task<ReviewDecisionDocument> DecideReviewAsync(
        string databasePath,
        string comparisonId,
        string decision,
        string reviewer,
        string reason,
        ReviewAssessment? assessment = null)
    {
        var arguments = new List<string>
        {
            "review-decide",
            "--db", Path.GetFullPath(databasePath),
            "--comparison-id", comparisonId,
            "--decision", decision,
            "--reviewer", reviewer,
            "--reason", reason,
        };
        if (assessment is not null)
        {
            arguments.AddRange(
            [
                "--study-comparability",
                assessment.StudyComparabilityStatus,
                "--study-confounding",
                assessment.StudyConfoundingStatus,
                "--comparison-validity",
                assessment.ComparisonValidityStatus,
                "--comparison-confounding",
                assessment.ComparisonConfoundingStatus,
                "--matching-basis",
                assessment.MatchingBasis,
            ]);
        }
        var result = await RunCliAsync(arguments.ToArray());
        return ReviewDecisionDocument.Load(result.StandardOutput);
    }

    internal async Task<ConceptCandidateListDocument>
        ConceptCandidatesAsync(
            string databasePath,
            string candidateKind,
            string query)
    {
        var arguments = new List<string>
        {
            "concept-candidates",
            "--db", Path.GetFullPath(databasePath),
            "--status", "OPEN",
            "--limit", "10000",
        };
        if (!string.IsNullOrWhiteSpace(candidateKind))
            arguments.AddRange(["--kind", candidateKind.Trim()]);
        if (!string.IsNullOrWhiteSpace(query))
            arguments.AddRange(["--query", query.Trim()]);
        var result = await RunCliAsync(arguments.ToArray());
        return ConceptCandidateListDocument.Load(
            result.StandardOutput,
            candidateKind.Trim(),
            query.Trim());
    }

    internal async Task<CanonicalConceptListDocument> ConceptsAsync(
        string databasePath,
        string conceptKind,
        string query)
    {
        if (string.IsNullOrWhiteSpace(conceptKind))
            throw new ArgumentException(
                "Concept kind is required.",
                nameof(conceptKind));
        var arguments = new List<string>
        {
            "concept-list",
            "--db", Path.GetFullPath(databasePath),
            "--status", "ACTIVE",
            "--kind", conceptKind.Trim(),
            "--limit", "10000",
        };
        if (!string.IsNullOrWhiteSpace(query))
            arguments.AddRange(["--query", query.Trim()]);
        var result = await RunCliAsync(arguments.ToArray());
        return CanonicalConceptListDocument.Load(
            result.StandardOutput,
            conceptKind.Trim(),
            query.Trim());
    }

    internal async Task<ConceptResolutionDocument> ResolveConceptAsync(
        string databasePath,
        string candidateUid,
        string action,
        string reviewer,
        string note,
        string canonicalName = "",
        string conceptUid = "",
        string alias = "")
    {
        if (string.IsNullOrWhiteSpace(candidateUid))
            throw new ArgumentException(
                "Candidate UID is required.",
                nameof(candidateUid));
        if (string.IsNullOrWhiteSpace(reviewer))
            throw new ArgumentException(
                "Reviewer is required.",
                nameof(reviewer));
        if (string.IsNullOrWhiteSpace(note))
            throw new ArgumentException(
                "Note is required.",
                nameof(note));

        var normalizedAction = action.Trim().ToUpperInvariant();
        var arguments = new List<string>
        {
            "concept-resolve",
            "--db", Path.GetFullPath(databasePath),
            "--candidate-uid", candidateUid.Trim(),
            "--action", normalizedAction,
            "--reviewer", reviewer.Trim(),
            "--note", note.Trim(),
        };
        switch (normalizedAction)
        {
            case "CREATE":
                if (string.IsNullOrWhiteSpace(canonicalName)
                    || string.IsNullOrWhiteSpace(alias))
                    throw new ArgumentException(
                        "CREATE requires a canonical name and alias.");
                arguments.AddRange(
                [
                    "--canonical-name", canonicalName.Trim(),
                    "--alias", alias.Trim(),
                ]);
                break;
            case "MERGE":
                if (string.IsNullOrWhiteSpace(conceptUid)
                    || string.IsNullOrWhiteSpace(alias))
                    throw new ArgumentException(
                        "MERGE requires a target concept UID and alias.");
                arguments.AddRange(
                [
                    "--concept-uid", conceptUid.Trim(),
                    "--alias", alias.Trim(),
                ]);
                break;
            case "REJECT":
                break;
            default:
                throw new ArgumentOutOfRangeException(
                    nameof(action),
                    action,
                    "Action must be CREATE, MERGE, or REJECT.");
        }
        var result = await RunCliAsync(arguments.ToArray());
        var resolution =
            ConceptResolutionDocument.Load(result.StandardOutput);
        if (!string.Equals(
                resolution.Candidate.CandidateUid,
                candidateUid.Trim(),
                StringComparison.Ordinal)
            || !string.Equals(
                resolution.Action,
                normalizedAction,
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "Concept resolution response does not match the request.");
        return resolution;
    }

    private Task<CliResult> RunCliAsync(params string[] arguments) =>
        RunCliAsync(null, arguments);

    private async Task<CliResult> RunCliAsync(
        IProgress<IngestProgressEvent>? progress,
        params string[] arguments)
    {
        var executable = _paths.PythonExecutable;
        var info = new ProcessStartInfo(executable)
        {
            WorkingDirectory = _serviceDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        info.Environment["PYTHONUTF8"] = "1";
        info.Environment["PYTHONIOENCODING"] = "utf-8";
        info.ArgumentList.Add(Path.Combine(
            _serviceDirectory,
            "inference_data_ai_cli.py"));
        foreach (var argument in arguments) info.ArgumentList.Add(argument);

        using var process = Process.Start(info)
            ?? throw new InvalidOperationException(
                "근거 DB Python 실행을 시작하지 못했습니다.");
        var outputTask = ReadCliStreamAsync(
            process.StandardOutput,
            false,
            progress);
        var errorTask = ReadCliStreamAsync(
            process.StandardError,
            true,
            progress);
        await process.WaitForExitAsync();
        var output = await outputTask;
        var error = await errorTask;
        if (process.ExitCode != 0)
        {
            var message = string.IsNullOrWhiteSpace(error) ? output : error;
            throw new InvalidOperationException(
                $"근거 DB 명령이 실패했습니다 ({process.ExitCode}): "
                + message.Trim());
        }
        return new CliResult(output, error);
    }

    private static async Task<string> ReadCliStreamAsync(
        StreamReader reader,
        bool parseProgress,
        IProgress<IngestProgressEvent>? progress)
    {
        var text = new StringBuilder();
        while (await reader.ReadLineAsync() is { } line)
        {
            if (parseProgress
                && line.StartsWith(
                    "PROGRESS_JSON ",
                    StringComparison.Ordinal))
            {
                try
                {
                    progress?.Report(IngestProgressEvent.Load(
                        line["PROGRESS_JSON ".Length..]));
                    continue;
                }
                catch (JsonException)
                {
                    // Preserve malformed telemetry as diagnostic stderr.
                }
            }
            if (parseProgress
                && line.StartsWith(
                    "AUTH_DIALOG_JSON ",
                    StringComparison.Ordinal))
            {
                progress?.Report(new IngestProgressEvent(
                    "AUTH_DIALOG",
                    "WAITING",
                    line["AUTH_DIALOG_JSON ".Length..],
                    string.Empty,
                    DateTime.UtcNow.ToString("O")));
                continue;
            }
            text.AppendLine(line);
        }
        return text.ToString();
    }

    private static string GetString(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value)
            ? value.ToString()
            : string.Empty;

    private static int GetInt32(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value)
        && value.TryGetInt32(out var result)
            ? result
            : 0;

    private sealed record CliResult(
        string StandardOutput,
        string StandardError);
}

internal sealed record EvidenceCitationRow(
    string EvidenceId,
    string SourcePath,
    string Sheet,
    string Range,
    string VerificationStatus)
{
    internal string FileName => Path.GetFileName(SourcePath);
}

internal sealed record EvidenceAnswerSession(
    string AnswerPath,
    string MarkdownPath,
    string SchemaVersion,
    string AnswerStatus,
    string Confidence,
    int RelevantStudyCount,
    int EligibleEffectCount,
    string Markdown,
    IReadOnlyList<EvidenceCitationRow> Citations,
    string RawJson)
{
    internal bool IsContextual => string.Equals(
        SchemaVersion,
        "table-first-context-answer-v1",
        StringComparison.Ordinal);

    internal bool IsRelevance => string.Equals(
        SchemaVersion,
        "table-first-relevance-result-v1",
        StringComparison.Ordinal);

    internal static EvidenceAnswerSession Load(
        string answerPath,
        string markdownPath)
    {
        var rawJson = File.ReadAllText(answerPath, Encoding.UTF8);
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        var coverage = root.GetProperty("coverage");
        var citationSource = root.TryGetProperty(
            "relatedCitations",
            out var relatedCitations)
            ? relatedCitations
            : root.GetProperty("citations");
        var citations = citationSource
            .EnumerateArray()
            .Select(item => new EvidenceCitationRow(
                item.GetProperty("evidenceId").GetString() ?? string.Empty,
                item.GetProperty("sourcePath").GetString() ?? string.Empty,
                item.GetProperty("sheet").GetString() ?? string.Empty,
                item.GetProperty("range").GetString() ?? string.Empty,
                item.GetProperty("verificationStatus").GetString()
                    ?? string.Empty))
            .ToList();
        return new EvidenceAnswerSession(
            Path.GetFullPath(answerPath),
            Path.GetFullPath(markdownPath),
            root.GetProperty("schemaVersion").GetString() ?? string.Empty,
            root.GetProperty("answerStatus").GetString() ?? string.Empty,
            root.TryGetProperty("confidence", out var confidence)
                ? confidence.GetString() ?? string.Empty
                : string.Empty,
            coverage.GetProperty("relevantStudyCount").GetInt32(),
            coverage.GetProperty("eligibleEffectCount").GetInt32(),
            File.ReadAllText(markdownPath, Encoding.UTF8),
            citations,
            rawJson);
    }
}

internal sealed record EvidenceDetailDocument(
    string EvidenceId,
    string TrustStatus,
    bool Trusted,
    string SourcePath,
    string Sheet,
    string Range,
    string RawJson)
{
    internal static EvidenceDetailDocument Load(string path)
    {
        var rawJson = File.ReadAllText(path, Encoding.UTF8);
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        var trust = root.GetProperty("trust");
        var source = root.GetProperty("source");
        var evidence = root.GetProperty("evidence");
        return new EvidenceDetailDocument(
            root.GetProperty("publicEvidenceId").GetString() ?? string.Empty,
            trust.GetProperty("status").GetString() ?? string.Empty,
            trust.GetProperty("trusted").GetBoolean(),
            source.GetProperty("sourcePath").GetString() ?? string.Empty,
            evidence.GetProperty("sheet").GetString() ?? string.Empty,
            evidence.GetProperty("range").GetString() ?? string.Empty,
            rawJson);
    }
}

internal sealed record IngestWorkbookResult(
    string Status,
    string WorkbookStatus,
    string RevisionUid,
    string SourcePath,
    string PublicAnalysisId,
    string JournalPath,
    string ManifestPath,
    string ArtifactDirectory,
    int StudyCount,
    bool? IntegrityOk,
    string RawJson);

internal sealed record IngestCorpusResult(
    string Status,
    string SourceRoot,
    string JournalPath,
    string ResultPath,
    int SelectedCount,
    int CompletedCount,
    int FailedCount,
    string RawJson);

internal sealed record FormPipelineCompleteResult(
    string Status,
    string ResultPath,
    string ReportPath,
    string ReviewPath,
    string ManifestPath,
    int PreflightTotalCount,
    int KnownFormCount,
    int ExcludedFormCount,
    int CaptureFailedCount,
    int FormGroupCount,
    int PendingFormGroupCount,
    int ApprovedFormGroupCount,
    int AnalysisErrorCount,
    int CorpusSelectedCount,
    int CorpusAttemptedCount,
    int CorpusCompletedCount,
    int CorpusFailedCount,
    string RawJson)
{
    internal bool IsComplete =>
        string.Equals(
            Status,
            "COMPLETED",
            StringComparison.OrdinalIgnoreCase)
        && PendingFormGroupCount == 0
        && AnalysisErrorCount == 0
        && CorpusFailedCount == 0;
}

internal sealed record FormPreflightDocument(
    string ReportPath,
    string Status,
    string SourceRoot,
    string KnownFormManifestPath,
    int TotalCount,
    int KnownFormCount,
    int SimilarReviewCount,
    int NewFormCount,
    int ExcludedFormCount,
    int CaptureFailedCount,
    IReadOnlyList<FormPreflightRow> Items)
{
    internal bool HasBlockingItems =>
        SimilarReviewCount > 0
        || NewFormCount > 0
        || ExcludedFormCount > 0
        || CaptureFailedCount > 0;

    internal static FormPreflightDocument Load(string path)
    {
        string rawJson;
        using (var stream = new FileStream(
                   path,
                   FileMode.Open,
                   FileAccess.Read,
                   FileShare.ReadWrite | FileShare.Delete))
        using (var reader = new StreamReader(
                   stream,
                   Encoding.UTF8,
                   detectEncodingFromByteOrderMarks: true))
            rawJson = reader.ReadToEnd();
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        if (!string.Equals(
                root.GetProperty("schemaVersion").GetString(),
                "excel-form-preflight-v1",
                StringComparison.Ordinal))
        {
            throw new InvalidDataException(
                "지원하지 않는 Excel 양식 사전 분석 결과입니다.");
        }
        var summary = root.GetProperty("summary");
        var rows = root.GetProperty("items")
            .EnumerateArray()
            .Select(item => new FormPreflightRow(
                item.GetProperty("status").GetString()
                    ?? string.Empty,
                item.GetProperty("fileName").GetString()
                    ?? string.Empty,
                item.GetProperty("sourcePath").GetString()
                    ?? string.Empty,
                item.GetProperty("contentSha256").GetString()
                    ?? string.Empty,
                item.GetProperty("captureAction").GetString()
                    ?? string.Empty,
                item.GetProperty("formSignatureId").GetString()
                    ?? string.Empty,
                item.TryGetProperty(
                        "similarity",
                        out var similarity)
                    && similarity.TryGetDouble(out var score)
                        ? score
                        : 0,
                item.GetProperty("nearestKnownSource").GetString()
                    ?? string.Empty,
                item.GetProperty("reason").GetString()
                    ?? string.Empty))
            .OrderBy(row => row.StatusSort)
            .ThenBy(
                row => row.FileName,
                StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return new FormPreflightDocument(
            Path.GetFullPath(path),
            root.GetProperty("status").GetString()
                ?? string.Empty,
            root.GetProperty("sourceRoot").GetString()
                ?? string.Empty,
            root.GetProperty("knownFormManifestPath").GetString()
                ?? string.Empty,
            summary.GetProperty("total").GetInt32(),
            summary.GetProperty("knownForms").GetInt32(),
            summary.GetProperty("similarReview").GetInt32(),
            summary.GetProperty("newForms").GetInt32(),
            summary.TryGetProperty(
                    "excludedForms",
                    out var excludedForms)
                && excludedForms.TryGetInt32(out var excludedCount)
                    ? excludedCount
                    : 0,
            summary.GetProperty("captureFailed").GetInt32(),
            rows);
    }
}

internal sealed record FormGroupReviewDocument(
    string ReviewPath,
    string PreflightStatus,
    int GroupCount,
    int PendingCount,
    int ApprovedCount,
    int ExcludedCount,
    int WorkbookCount,
    IReadOnlyList<FormFamilyGroupRow> Groups)
{
    internal static FormGroupReviewDocument Load(string path)
    {
        var rawJson = File.ReadAllText(path, Encoding.UTF8);
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        if (!string.Equals(
                root.GetProperty("schemaVersion").GetString(),
                "excel-form-group-review-v1",
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "지원하지 않는 Excel 양식군 검토 결과입니다.");
        var summary = root.GetProperty("summary");
        var groups = root.GetProperty("groups")
            .EnumerateArray()
            .Select(item =>
            {
                var candidateStatuses = item.TryGetProperty(
                        "candidateStatuses",
                        out var statuses)
                    && statuses.ValueKind == JsonValueKind.Object
                        ? string.Join(
                            " · ",
                            statuses.EnumerateObject().Select(
                                property =>
                                    $"{property.Name} "
                                    + property.Value.GetInt32()
                                        .ToString("N0")))
                        : string.Empty;
                var samples = item.TryGetProperty(
                        "sampleSources",
                        out var sampleSources)
                    && sampleSources.ValueKind == JsonValueKind.Array
                        ? sampleSources.EnumerateArray()
                            .Select(value => value.GetString()
                                ?? string.Empty)
                            .Where(value => value.Length > 0)
                            .ToArray()
                        : [];
                return new FormFamilyGroupRow(
                    item.GetProperty("familyId").GetString()
                        ?? string.Empty,
                    item.GetProperty("displayName").GetString()
                        ?? string.Empty,
                    item.GetProperty("decisionStatus").GetString()
                        ?? string.Empty,
                    item.GetProperty("memberCount").GetInt32(),
                    item.GetProperty("representativeSource").GetString()
                        ?? string.Empty,
                    samples,
                    item.TryGetProperty(
                            "averageSimilarity",
                            out var averageSimilarity)
                        && averageSimilarity.TryGetDouble(
                            out var similarity)
                            ? similarity
                            : 0,
                    item.GetProperty(
                            "nearestKnownFormSignatureId")
                        .GetString()
                        ?? string.Empty,
                    item.GetProperty("nearestKnownSource").GetString()
                        ?? string.Empty,
                    candidateStatuses,
                    item.GetProperty("contractAvailable").GetBoolean(),
                    item.GetProperty("validationStatus").GetString()
                        ?? string.Empty,
                    item.GetProperty("validationSampleCount")
                        .GetInt32(),
                    item.GetProperty("reviewer").GetString()
                        ?? string.Empty,
                    item.GetProperty("notes").GetString()
                        ?? string.Empty);
            })
            .ToArray();
        return new FormGroupReviewDocument(
            Path.GetFullPath(path),
            root.GetProperty("preflightStatus").GetString()
                ?? string.Empty,
            summary.GetProperty("groupCount").GetInt32(),
            summary.GetProperty("pendingCount").GetInt32(),
            summary.GetProperty("approvedCount").GetInt32(),
            summary.GetProperty("excludedCount").GetInt32(),
            summary.GetProperty("workbookCount").GetInt32(),
            groups);
    }
}

internal sealed record FormFamilyGroupRow(
    string FamilyId,
    string DisplayName,
    string DecisionStatus,
    int MemberCount,
    string RepresentativeSource,
    IReadOnlyList<string> SampleSources,
    double AverageSimilarity,
    string NearestKnownFormSignatureId,
    string NearestKnownSource,
    string CandidateStatuses,
    bool ContractAvailable,
    string ValidationStatus,
    int ValidationSampleCount,
    string Reviewer,
    string Notes)
{
    public string DecisionStatusDisplay => DecisionStatus switch
    {
        "PENDING" => "AI 분석 대기",
        "ANALYZED_PENDING_APPROVAL" => "사람 승인 대기",
        "APPROVED_NEW" => "신규 양식 등록",
        "LINKED_EXISTING" => "기존 양식 연결",
        "EXCLUDED" => "전체 처리 제외",
        _ => DecisionStatus,
    };

    public string RepresentativeFile =>
        Path.GetFileName(RepresentativeSource);

    public string SampleDisplay =>
        string.Join(
            " · ",
            SampleSources.Select(Path.GetFileName));

    public string SimilarityDisplay =>
        AverageSimilarity.ToString("P0");

    public string ValidationDisplay => !ContractAvailable
        ? "미분석"
        : $"{ValidationStatus} · {ValidationSampleCount:N0}개";

    public string NearestKnownFile =>
        Path.GetFileName(NearestKnownSource);

    public bool CanRegisterNew =>
        ContractAvailable
        && string.Equals(
            ValidationStatus,
            "PASSED",
            StringComparison.Ordinal);
}

internal sealed record FormPreflightRow(
    string Status,
    string FileName,
    string SourcePath,
    string ContentSha256,
    string CaptureAction,
    string FormSignatureId,
    double Similarity,
    string NearestKnownSource,
    string Reason)
{
    public string StatusDisplay => Status switch
    {
        "KNOWN_FORM" => "기존 양식",
        "SIMILAR_FORM_REVIEW" => "유사 양식 · 검토",
        "NEW_FORM" => "신규 양식 · 보류",
        "EXCLUDED_FORM" => "사람 판정 · 제외",
        "CAPTURE_FAILED" => "COM 추출 실패",
        "RUNNING" => "COM 분석 중",
        "PENDING" => "분석 전",
        _ => Status,
    };

    public string SimilarityDisplay =>
        Status is "PENDING" or "RUNNING" or "CAPTURE_FAILED"
            ? "-"
            : Similarity.ToString("P0");

    public string NearestKnownFile =>
        Path.GetFileName(NearestKnownSource);

    internal int StatusSort => Status switch
    {
        "CAPTURE_FAILED" => 0,
        "EXCLUDED_FORM" => 1,
        "NEW_FORM" => 2,
        "SIMILAR_FORM_REVIEW" => 3,
        "KNOWN_FORM" => 4,
        _ => 4,
    };
}

internal sealed record IngestProgressEvent(
    string Stage,
    string Status,
    string Detail,
    string SourcePath,
    string Timestamp)
{
    internal static IngestProgressEvent Load(string rawJson)
    {
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        return new IngestProgressEvent(
            Read(root, "stage"),
            Read(root, "status"),
            Read(root, "detail"),
            Read(root, "sourcePath"),
            Read(root, "timestamp"));
    }

    private static string Read(JsonElement root, string name) =>
        root.TryGetProperty(name, out var value)
            ? value.ToString()
            : string.Empty;
}

internal sealed record RelatedStudyRow(
    string PublicDataId,
    string Title,
    double SimilarityScore,
    string SourcePath);

internal sealed record RelatedStudiesDocument(
    string TargetIdentifier,
    int ExactDuplicateSourceCount,
    IReadOnlyList<RelatedStudyRow> Studies)
{
    internal static RelatedStudiesDocument Load(string path)
    {
        using var document = JsonDocument.Parse(
            File.ReadAllText(path, Encoding.UTF8));
        var root = document.RootElement;
        var summary = root.GetProperty("summary");
        var studies = root.GetProperty("relatedStudies")
            .EnumerateArray()
            .Select(item => new RelatedStudyRow(
                item.GetProperty("publicDataId").GetString()
                    ?? string.Empty,
                item.GetProperty("title").GetString()
                    ?? string.Empty,
                item.GetProperty("similarityScore").GetDouble(),
                item.GetProperty("source")
                    .GetProperty("sourcePath")
                    .GetString()
                    ?? string.Empty))
            .ToList();
        return new RelatedStudiesDocument(
            root.GetProperty("targetIdentifier").GetString()
                ?? string.Empty,
            summary.GetProperty("exactContentDuplicateSourceCount")
                .GetInt32(),
            studies);
    }
}

internal sealed record ReviewQueueRow(
    string PublicComparisonId,
    string PublicDataId,
    string StudyTitle,
    string FileName,
    string SourcePath,
    string ComparisonValidityStatus,
    string MatchingBasis);

internal sealed record ReviewQueueDocument(
    IReadOnlyList<ReviewQueueRow> Items)
{
    internal static ReviewQueueDocument Load(string path)
    {
        using var document = JsonDocument.Parse(
            File.ReadAllText(path, Encoding.UTF8));
        var items = document.RootElement.GetProperty("items")
            .EnumerateArray()
            .Select(item => new ReviewQueueRow(
                item.GetProperty("publicComparisonId").GetString()
                    ?? string.Empty,
                item.GetProperty("publicDataId").GetString()
                    ?? string.Empty,
                item.GetProperty("studyTitle").GetString()
                    ?? string.Empty,
                item.GetProperty("fileName").GetString()
                    ?? string.Empty,
                item.GetProperty("sourcePath").GetString()
                    ?? string.Empty,
                item.GetProperty("comparisonValidityStatus").GetString()
                    ?? string.Empty,
                item.GetProperty("matchingBasis").GetString()
                    ?? string.Empty))
            .ToList();
        return new ReviewQueueDocument(items);
    }
}

internal sealed record ReviewEvidenceRow(
    string EvidenceId,
    string Role,
    string SourcePath,
    string Sheet,
    string Range,
    string ValueSummary);

internal sealed record ReviewDetailDocument(
    string PublicComparisonId,
    string PublicDataId,
    string SourcePath,
    string StudyTitle,
    string StudyComparabilityStatus,
    string StudyConfoundingStatus,
    string ComparisonValidityStatus,
    string ComparisonConfoundingStatus,
    string MatchingBasis,
    string ComparedArmLabel,
    string ControlArmLabel,
    bool ApprovalReady,
    string BlockerSummary,
    IReadOnlyList<ReviewEvidenceRow> Evidence,
    string RawJson)
{
    internal static ReviewDetailDocument Load(string path)
    {
        var rawJson = File.ReadAllText(path, Encoding.UTF8);
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        var source = root.GetProperty("source");
        var study = root.GetProperty("study");
        var comparison = root.GetProperty("comparison");
        var readiness = root.GetProperty("approvalReadiness");
        var sourcePath = source.GetProperty("path").GetString()
            ?? string.Empty;
        var evidence = new List<ReviewEvidenceRow>();
        foreach (var item in comparison.GetProperty("evidence")
                     .EnumerateArray())
        {
            evidence.Add(EvidenceFromJson(
                item,
                "COMPARISON",
                sourcePath,
                string.Empty));
        }
        foreach (var pair in root.GetProperty("pairedObservations")
                     .EnumerateArray())
        {
            var outcome = pair.GetProperty("outcomeLabel").GetString()
                ?? pair.GetProperty("outcomeKey").GetString()
                ?? string.Empty;
            var unit = pair.GetProperty("outcomeUnit").GetString()
                ?? string.Empty;
            AddObservationEvidence(
                evidence,
                pair.GetProperty("comparedObservation"),
                "TEST",
                sourcePath,
                outcome,
                unit);
            AddObservationEvidence(
                evidence,
                pair.GetProperty("controlObservation"),
                "CONTROL",
                sourcePath,
                outcome,
                unit);
        }
        var blockers = readiness.GetProperty("blockers")
            .EnumerateArray()
            .Select(item =>
                $"{item.GetProperty("code").GetString()}: "
                + item.GetProperty("message").GetString())
            .ToList();
        return new ReviewDetailDocument(
            root.GetProperty("publicComparisonId").GetString()
                ?? string.Empty,
            root.GetProperty("publicDataId").GetString()
                ?? string.Empty,
            sourcePath,
            study.GetProperty("title").GetString() ?? string.Empty,
            study.GetProperty("comparabilityStatus").GetString()
                ?? string.Empty,
            study.GetProperty("confoundingStatus").GetString()
                ?? string.Empty,
            comparison.GetProperty("validityStatus").GetString()
                ?? string.Empty,
            comparison.GetProperty("confoundingStatus").GetString()
                ?? string.Empty,
            comparison.GetProperty("matchingBasis").GetString()
                ?? string.Empty,
            comparison.GetProperty("comparedArm")
                .GetProperty("label").GetString()
                ?? string.Empty,
            comparison.GetProperty("controlArm")
                .GetProperty("label").GetString()
                ?? string.Empty,
            readiness.GetProperty("ready").GetBoolean(),
            string.Join("\n", blockers),
            evidence
                .GroupBy(item => (
                    item.EvidenceId,
                    item.Role,
                    item.ValueSummary))
                .Select(group => group.First())
                .ToList(),
            rawJson);
    }

    private static void AddObservationEvidence(
        ICollection<ReviewEvidenceRow> rows,
        JsonElement observation,
        string role,
        string sourcePath,
        string outcome,
        string unit)
    {
        var value = ValueSummary(observation, outcome, unit);
        foreach (var item in observation.GetProperty("evidence")
                     .EnumerateArray())
            rows.Add(EvidenceFromJson(item, role, sourcePath, value));
    }

    private static ReviewEvidenceRow EvidenceFromJson(
        JsonElement item,
        string role,
        string sourcePath,
        string value) =>
        new(
            item.GetProperty("publicEvidenceId").GetString()
                ?? string.Empty,
            role,
            sourcePath,
            item.GetProperty("sheet").GetString() ?? string.Empty,
            item.GetProperty("range").GetString() ?? string.Empty,
            value);

    private static string ValueSummary(
        JsonElement observation,
        string outcome,
        string unit)
    {
        static string JsonText(JsonElement element, string name)
        {
            if (!element.TryGetProperty(name, out var value)
                || value.ValueKind is JsonValueKind.Null
                    or JsonValueKind.Undefined)
                return string.Empty;
            return value.ValueKind == JsonValueKind.String
                ? value.GetString() ?? string.Empty
                : value.GetRawText();
        }
        var parts = new List<string>();
        var valueNumber = JsonText(observation, "valueNumber");
        var valueText = JsonText(observation, "valueText");
        if (valueNumber.Length > 0)
            parts.Add($"값 {valueNumber}{(unit.Length > 0 ? " " + unit : "")}");
        else if (valueText.Length > 0)
            parts.Add($"값 {valueText}");
        var numerator = JsonText(observation, "numerator");
        var denominator = JsonText(observation, "denominator");
        if (numerator.Length > 0 || denominator.Length > 0)
            parts.Add($"{numerator}/{denominator}");
        var sampleSize = JsonText(observation, "sampleSize");
        if (sampleSize.Length > 0) parts.Add($"n={sampleSize}");
        return $"{outcome}: {string.Join(", ", parts)}";
    }
}

internal sealed record ReviewAssessment(
    string StudyComparabilityStatus,
    string StudyConfoundingStatus,
    string ComparisonValidityStatus,
    string ComparisonConfoundingStatus,
    string MatchingBasis);

internal sealed record ReviewDecisionDocument(
    string PublicComparisonId,
    string PublicDataId,
    string Decision,
    bool AggregationEligible,
    IReadOnlyList<string> EffectPublicIds,
    string RawJson)
{
    internal static ReviewDecisionDocument Load(string rawJson)
    {
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        return new ReviewDecisionDocument(
            root.GetProperty("publicComparisonId").GetString()
                ?? string.Empty,
            root.GetProperty("publicDataId").GetString()
                ?? string.Empty,
            root.GetProperty("decision").GetString() ?? string.Empty,
            root.GetProperty("comparisonAggregationEligible")
                .GetBoolean(),
            root.GetProperty("effectPublicIds")
                .EnumerateArray()
                .Select(item => item.GetString() ?? string.Empty)
                .Where(item => item.Length > 0)
                .ToList(),
            rawJson);
    }
}

internal sealed record ConceptCandidateRow(
    int SchemaCandidateId,
    string CandidateUid,
    string CandidateKind,
    string NormalizedValue,
    string OriginalValue,
    string SuggestedCanonicalName,
    int OccurrenceCount,
    string Status,
    string FirstSeenAt,
    string LastSeenAt)
{
    internal bool IsConceptCandidate =>
        CandidateKind.StartsWith(
            "CONCEPT:",
            StringComparison.OrdinalIgnoreCase)
        && ConceptKind.Length > 0;

    internal string ConceptKind =>
        CandidateKind.StartsWith(
            "CONCEPT:",
            StringComparison.OrdinalIgnoreCase)
            ? CandidateKind["CONCEPT:".Length..].Trim().ToUpperInvariant()
            : string.Empty;
}

internal sealed record ConceptCandidateListFilters(
    string Status,
    string CandidateKind,
    string Query,
    int Limit);

internal sealed record ConceptCandidateListDocument(
    string SchemaVersion,
    ConceptCandidateListFilters Filters,
    int Count,
    IReadOnlyList<ConceptCandidateRow> Candidates,
    string RawJson)
{
    internal static ConceptCandidateListDocument Load(
        string rawJson,
        string requestedCandidateKind,
        string requestedQuery)
    {
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        const string expectedSchema = "concept-candidate-list-v1";
        var schemaVersion = RequiredString(root, "schemaVersion");
        if (!string.Equals(
                schemaVersion,
                expectedSchema,
                StringComparison.Ordinal))
            throw new InvalidDataException(
                $"Unsupported concept candidate schemaVersion: "
                + $"'{schemaVersion}'. Expected '{expectedSchema}'.");
        var filterElement = root.GetProperty("filters");
        var filters = new ConceptCandidateListFilters(
            RequiredString(filterElement, "status"),
            RequiredString(filterElement, "candidateKind"),
            RequiredString(filterElement, "query"),
            filterElement.GetProperty("limit").GetInt32());
        if (!string.Equals(
                filters.Status,
                "OPEN",
                StringComparison.Ordinal)
            || filters.Limit != 10000
            || !string.Equals(
                filters.CandidateKind,
                requestedCandidateKind.Trim().ToUpperInvariant(),
                StringComparison.Ordinal)
            || !string.Equals(
                filters.Query,
                requestedQuery.Trim(),
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "Concept candidate response filters do not match "
                + "the request.");
        var candidates = root.GetProperty("candidates")
            .EnumerateArray()
            .Select(item => new ConceptCandidateRow(
                item.GetProperty("schemaCandidateId").GetInt32(),
                RequiredString(item, "candidateUid"),
                RequiredString(item, "candidateKind"),
                RequiredString(item, "normalizedValue"),
                RequiredString(item, "originalValue"),
                RequiredString(item, "suggestedCanonicalName"),
                item.GetProperty("occurrenceCount").GetInt32(),
                RequiredString(item, "status"),
                RequiredString(item, "firstSeenAt"),
                RequiredString(item, "lastSeenAt")))
            .ToList();
        var count = root.GetProperty("count").GetInt32();
        if (count != candidates.Count)
            throw new InvalidDataException(
                "Concept candidate count does not match the payload.");
        if (candidates.Any(item =>
                !string.Equals(
                    item.Status,
                    "OPEN",
                    StringComparison.Ordinal)))
            throw new InvalidDataException(
                "The OPEN candidate response contained a non-OPEN row.");
        return new ConceptCandidateListDocument(
            schemaVersion,
            filters,
            count,
            candidates,
            rawJson);
    }

    private static string RequiredString(
        JsonElement element,
        string name) =>
        element.GetProperty(name).GetString()
        ?? throw new InvalidDataException(
            $"Concept candidate field '{name}' is null.");
}

internal sealed record ConceptAliasRow(
    int AliasId,
    string AliasUid,
    string AliasText,
    string NormalizedAlias,
    string Language,
    string Source,
    double Confidence,
    string CreatedAt);

internal sealed record CanonicalConceptRow(
    int ConceptId,
    string ConceptUid,
    string ConceptKind,
    string CanonicalName,
    string NormalizedName,
    string Description,
    string LifecycleStatus,
    string CreatedAt,
    string UpdatedAt,
    IReadOnlyList<ConceptAliasRow> Aliases)
{
    internal string AliasSummary => string.Join(
        ", ",
        Aliases.Select(item => item.AliasText));
}

internal sealed record CanonicalConceptListFilters(
    string ConceptKind,
    string LifecycleStatus,
    string Query,
    int Limit);

internal sealed record CanonicalConceptListDocument(
    string SchemaVersion,
    CanonicalConceptListFilters Filters,
    int Count,
    IReadOnlyList<CanonicalConceptRow> Concepts,
    string RawJson)
{
    internal static CanonicalConceptListDocument Load(
        string rawJson,
        string requestedConceptKind,
        string requestedQuery)
    {
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        const string expectedSchema = "canonical-concept-list-v1";
        var schemaVersion = RequiredString(root, "schemaVersion");
        if (!string.Equals(
                schemaVersion,
                expectedSchema,
                StringComparison.Ordinal))
            throw new InvalidDataException(
                $"Unsupported canonical concept schemaVersion: "
                + $"'{schemaVersion}'. Expected '{expectedSchema}'.");
        var filterElement = root.GetProperty("filters");
        var filters = new CanonicalConceptListFilters(
            RequiredString(filterElement, "conceptKind"),
            RequiredString(filterElement, "lifecycleStatus"),
            RequiredString(filterElement, "query"),
            filterElement.GetProperty("limit").GetInt32());
        if (!string.Equals(
                filters.LifecycleStatus,
                "ACTIVE",
                StringComparison.Ordinal)
            || filters.Limit != 10000
            || !string.Equals(
                filters.ConceptKind,
                requestedConceptKind.Trim().ToUpperInvariant(),
                StringComparison.Ordinal)
            || !string.Equals(
                filters.Query,
                requestedQuery.Trim(),
                StringComparison.Ordinal))
            throw new InvalidDataException(
                "Canonical concept response filters do not match "
                + "the request.");
        var concepts = root.GetProperty("concepts")
            .EnumerateArray()
            .Select(ParseConcept)
            .ToList();
        var count = root.GetProperty("count").GetInt32();
        if (count != concepts.Count)
            throw new InvalidDataException(
                "Canonical concept count does not match the payload.");
        if (concepts.Any(item =>
                !string.Equals(
                    item.LifecycleStatus,
                    "ACTIVE",
                    StringComparison.Ordinal)
                || !string.Equals(
                    item.ConceptKind,
                    requestedConceptKind,
                    StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException(
                "The ACTIVE same-kind concept response contained "
                + "an unexpected row.");
        return new CanonicalConceptListDocument(
            schemaVersion,
            filters,
            count,
            concepts,
            rawJson);
    }

    private static CanonicalConceptRow ParseConcept(JsonElement item)
    {
        var aliases = item.GetProperty("aliases")
            .EnumerateArray()
            .Select(alias => new ConceptAliasRow(
                alias.GetProperty("aliasId").GetInt32(),
                RequiredString(alias, "aliasUid"),
                RequiredString(alias, "aliasText"),
                RequiredString(alias, "normalizedAlias"),
                RequiredString(alias, "language"),
                RequiredString(alias, "source"),
                alias.GetProperty("confidence").GetDouble(),
                RequiredString(alias, "createdAt")))
            .ToList();
        return new CanonicalConceptRow(
            item.GetProperty("conceptId").GetInt32(),
            RequiredString(item, "conceptUid"),
            RequiredString(item, "conceptKind"),
            RequiredString(item, "canonicalName"),
            RequiredString(item, "normalizedName"),
            RequiredString(item, "description"),
            RequiredString(item, "lifecycleStatus"),
            RequiredString(item, "createdAt"),
            RequiredString(item, "updatedAt"),
            aliases);
    }

    private static string RequiredString(
        JsonElement element,
        string name) =>
        element.GetProperty(name).GetString()
        ?? throw new InvalidDataException(
            $"Canonical concept field '{name}' is null.");
}

internal sealed record ConceptResolutionCandidate(
    int SchemaCandidateId,
    string CandidateUid,
    string CandidateKind,
    string NormalizedValue,
    string OriginalValue,
    string SuggestedCanonicalName,
    string Status);

internal sealed record ConceptResolutionConcept(
    int ConceptId,
    string ConceptUid,
    string ConceptKind,
    string CanonicalName,
    string NormalizedName);

internal sealed record ConceptResolutionAlias(
    int AliasId,
    string AliasUid,
    string AliasText,
    string NormalizedAlias,
    string Source,
    double Confidence);

internal sealed record ConceptResolutionDocument(
    string SchemaVersion,
    string ResolutionUid,
    ConceptResolutionCandidate Candidate,
    string Action,
    ConceptResolutionConcept? Concept,
    ConceptResolutionAlias? Alias,
    string Reviewer,
    string Note,
    string ResolvedAt,
    bool IdempotentReplay,
    string RawJson)
{
    internal static ConceptResolutionDocument Load(string rawJson)
    {
        using var document = JsonDocument.Parse(rawJson);
        var root = document.RootElement;
        const string expectedSchema = "concept-resolution-v1";
        var schemaVersion = RequiredString(root, "schemaVersion");
        if (!string.Equals(
                schemaVersion,
                expectedSchema,
                StringComparison.Ordinal))
            throw new InvalidDataException(
                $"Unsupported concept resolution schemaVersion: "
                + $"'{schemaVersion}'. Expected '{expectedSchema}'.");
        var candidateElement = root.GetProperty("candidate");
        var candidate = new ConceptResolutionCandidate(
            candidateElement.GetProperty("schemaCandidateId").GetInt32(),
            RequiredString(candidateElement, "candidateUid"),
            RequiredString(candidateElement, "candidateKind"),
            RequiredString(candidateElement, "normalizedValue"),
            RequiredString(candidateElement, "originalValue"),
            RequiredString(candidateElement, "suggestedCanonicalName"),
            RequiredString(candidateElement, "status"));
        ConceptResolutionConcept? concept = null;
        if (root.TryGetProperty("concept", out var conceptElement)
            && conceptElement.ValueKind == JsonValueKind.Object)
            concept = new ConceptResolutionConcept(
                conceptElement.GetProperty("conceptId").GetInt32(),
                RequiredString(conceptElement, "conceptUid"),
                RequiredString(conceptElement, "conceptKind"),
                RequiredString(conceptElement, "canonicalName"),
                RequiredString(conceptElement, "normalizedName"));
        ConceptResolutionAlias? alias = null;
        if (root.TryGetProperty("alias", out var aliasElement)
            && aliasElement.ValueKind == JsonValueKind.Object)
            alias = new ConceptResolutionAlias(
                aliasElement.GetProperty("aliasId").GetInt32(),
                RequiredString(aliasElement, "aliasUid"),
                RequiredString(aliasElement, "aliasText"),
                RequiredString(aliasElement, "normalizedAlias"),
                RequiredString(aliasElement, "source"),
                aliasElement.GetProperty("confidence").GetDouble());
        return new ConceptResolutionDocument(
            schemaVersion,
            RequiredString(root, "resolutionUid"),
            candidate,
            RequiredString(root, "action"),
            concept,
            alias,
            RequiredString(root, "reviewer"),
            RequiredString(root, "note"),
            RequiredString(root, "resolvedAt"),
            root.GetProperty("idempotentReplay").GetBoolean(),
            rawJson);
    }

    private static string RequiredString(
        JsonElement element,
        string name) =>
        element.GetProperty(name).GetString()
        ?? throw new InvalidDataException(
            $"Concept resolution field '{name}' is null.");
}
