using JinoSupporter.Web.Services;

string root = Path.Combine(Path.GetTempPath(), "jino-bom-cache-verify-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
Environment.SetEnvironmentVariable("JINOSUPPORTER_DATA_DIR", root);

try
{
    var paths = new AppPathsService();
    var settings = new NgRateSettingsService(paths);
    var cache = new BmesBomCacheService(settings);
    var originalDate = new DateTime(2026, 7, 29);
    var firstQuery = new BmesBomCacheQuery("TIU-L5S3-01", "3200", originalDate, 6, 2000);

    cache.Save(
        firstQuery,
        [
            new BmesBomMaterialCandidate
            {
                ProductCode = "P-L",
                ProductName = "TIU-L5S3-01-L-ZZ",
                ParentMaterialCode = "P-L",
                ParentMaterialName = "TIU-L5S3-01-L-ZZ",
                MaterialCode = "R-S-071110900",
                MaterialName = "COIL-L5S3-01",
                BomLevel = 1,
                BomPath = "P-L>R-S-071110900",
                BomPathText = "TIU-L5S3-01-L-ZZ > COIL-L5S3-01",
                SourceRows = 1,
                Source = "BOM",
            },
            new BmesBomMaterialCandidate
            {
                ProductCode = "P-R",
                ProductName = "TIU-L5S3-01-R-ZZ",
                ParentMaterialCode = "P-R",
                ParentMaterialName = "TIU-L5S3-01-R-ZZ",
                MaterialCode = "R-S-046137400",
                MaterialName = "F-PCB-L5S3-01-R",
                UsageQty = 1.25m,
                UsageUnit = "EA",
                BomLevel = 1,
                BomPath = "P-R>R-S-046137400",
                BomPathText = "TIU-L5S3-01-R-ZZ > F-PCB-L5S3-01-R",
                SourceRows = 1,
                Source = "BOM",
            },
        ]);

    // WorkDate intentionally differs: model cache must still be reused until Server is forced.
    var nextDayQuery = firstQuery with { WorkDate = originalDate.AddDays(1) };
    BmesBomCacheEntry loaded = cache.TryLoad(nextDayQuery)
        ?? throw new InvalidOperationException("Cache miss.");

    if (loaded.Rows.Count != 2 ||
        loaded.Query.WorkDate != originalDate ||
        loaded.Rows[1].UsageQty != 1.25m ||
        loaded.Rows[0].MaterialName != "COIL-L5S3-01")
    {
        throw new InvalidOperationException("Cache round-trip mismatch.");
    }

    BmesBomCachedModel listed = cache.ListSuccessfulModels()
        .Single(m => m.ModelName == "TIU-L5S3-01");
    if (listed.RowCount != 2)
        throw new InvalidOperationException("Successful-model index mismatch.");

    cache.SaveModelCatalog(
        "3200",
        originalDate,
        [
            new BmesBomModelCandidate
            {
                ProductCode = "C-S-000001",
                ProductName = "ASSY REAR-TAPE-338",
            },
            new BmesBomModelCandidate
            {
                ProductCode = "C-S-000002",
                ProductName = "TIU-L5S3-01",
            },
        ]);

    BmesBomModelCatalogEntry catalog = cache.LoadModelCatalog("3200");
    if (catalog.Models.Count != 2 ||
        catalog.WorkDate != originalDate ||
        catalog.Models[0].ProductName != "ASSY REAR-TAPE-338" ||
        string.IsNullOrWhiteSpace(catalog.SyncedAt))
    {
        throw new InvalidOperationException("Model catalog round-trip mismatch.");
    }

    Console.WriteLine(
        $"PASS rows={loaded.Rows.Count} cachedDate={loaded.Query.WorkDate:yyyy-MM-dd} modelNames={catalog.Models.Count}");
}
finally
{
    Environment.SetEnvironmentVariable("JINOSUPPORTER_DATA_DIR", null);
    if (Directory.Exists(root))
    {
        try
        {
            Directory.Delete(root, recursive: true);
        }
        catch (IOException)
        {
            // Windows may hold a just-closed SQLite WAL handle until process exit.
        }
    }
}
