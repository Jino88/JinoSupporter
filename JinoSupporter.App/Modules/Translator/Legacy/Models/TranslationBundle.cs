namespace CustomKeyboardCSharp.Models;

public sealed class TranslationBundle
{
    public List<TranslationOption> Options { get; set; } = [];
    public List<GlossaryEntry> Glossary { get; set; } = [];
}
