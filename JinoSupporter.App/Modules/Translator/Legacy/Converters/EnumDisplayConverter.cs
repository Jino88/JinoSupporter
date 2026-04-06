using System.Globalization;
using System.Windows.Data;
using CustomKeyboardCSharp.Models;

namespace CustomKeyboardCSharp.Converters;

public sealed class EnumDisplayConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value switch
        {
            AiProvider.Gemini => "Gemini",
            AiProvider.OpenAi => "ChatGPT",
            TranslationDirection.AutoToVietnamese => "Auto -> Vietnamese",
            TranslationDirection.KoreanToVietnamese => "Korean -> Vietnamese",
            TranslationDirection.EnglishToVietnamese => "English -> Vietnamese",
            TranslationDirection.VietnameseToKorean => "Vietnamese -> Korean",
            _ => value?.ToString() ?? string.Empty
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => System.Windows.Data.Binding.DoNothing;
}
