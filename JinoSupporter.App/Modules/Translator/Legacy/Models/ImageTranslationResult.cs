namespace CustomKeyboardCSharp.Models;

public sealed class ImageTranslationResult
{
    public string DetectedText { get; set; } = string.Empty;
    public List<TranslationOption> Options { get; set; } = [];
    public List<GlossaryEntry> Glossary { get; set; } = [];

    public TranslationBundle ToBundle() => new()
    {
        Options = Options,
        Glossary = Glossary
    };
}
