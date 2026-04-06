namespace QuickShareClone.Server;

public sealed class UploadOptions
{
    public string RootPath { get; set; } = "App_Data/uploads";
    public int ChunkExpirationHours { get; set; } = 24;
}
