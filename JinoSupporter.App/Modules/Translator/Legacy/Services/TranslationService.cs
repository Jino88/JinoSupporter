using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Net.Http;
using CustomKeyboardCSharp.Models;

namespace CustomKeyboardCSharp.Services;

public sealed class TranslationService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public async Task<TranslationBundle> TranslateAsync(
        AppSettings settings,
        string text,
        TranslationDirection direction,
        CancellationToken cancellationToken)
    {
        var prompt = BuildPrompt(text, direction);
        var content = settings.SelectedProvider switch
        {
            AiProvider.Gemini => await RequestGeminiAsync(settings, prompt, cancellationToken),
            _ => await RequestOpenAiAsync(settings, prompt, cancellationToken)
        };

        var bundle = JsonSerializer.Deserialize<TranslationBundle>(content, JsonOptions)
            ?? throw new InvalidOperationException("The translation response was empty.");

        if (bundle.Options.Count == 0)
        {
            throw new InvalidOperationException("No translation options were returned.");
        }

        return bundle;
    }

    public async Task<ImageTranslationResult> TranslateImageAsync(
        AppSettings settings,
        byte[] imageBytes,
        TranslationDirection direction,
        CancellationToken cancellationToken)
    {
        var prompt = BuildImagePrompt(direction);
        var content = settings.SelectedProvider switch
        {
            AiProvider.Gemini => await RequestGeminiImageAsync(settings, prompt, imageBytes, cancellationToken),
            _ => await RequestOpenAiImageAsync(settings, prompt, imageBytes, cancellationToken)
        };

        var result = JsonSerializer.Deserialize<ImageTranslationResult>(content, JsonOptions)
            ?? throw new InvalidOperationException("The image translation response was empty.");

        if (result.Options.Count == 0)
        {
            throw new InvalidOperationException("No translation options were returned from the image.");
        }

        return result;
    }

    public bool HasValidApiKey(AppSettings settings)
    {
        var key = GetApiKey(settings).Trim();
        if (string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        return settings.SelectedProvider switch
        {
            AiProvider.Gemini => key.StartsWith("AIza", StringComparison.Ordinal),
            _ => key.StartsWith("sk-", StringComparison.Ordinal) || key.StartsWith("sk-proj-", StringComparison.Ordinal)
        };
    }

    private async Task<string> RequestOpenAiAsync(AppSettings settings, string prompt, CancellationToken cancellationToken)
    {
        using var client = CreateClient(settings);
        var payload = new
        {
            model = "gpt-4o-mini",
            messages = new object[]
            {
                new
                {
                    role = "system",
                    content = "You are a concise translation assistant. Always return strict JSON. Keep translated text in the requested target language. Keep glossary words in the source language. Keep glossary meanings in the target language. Write nuance explanations in Korean."
                },
                new { role = "user", content = prompt }
            }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", GetApiKey(settings));
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using var response = await client.SendAsync(request, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();

        using var document = JsonDocument.Parse(body);
        return document.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString() ?? throw new InvalidOperationException("OpenAI returned no content.");
    }

    private async Task<string> RequestGeminiAsync(AppSettings settings, string prompt, CancellationToken cancellationToken)
    {
        using var client = CreateClient(settings);
        var payload = new
        {
            contents = new[]
            {
                new
                {
                    parts = new[]
                    {
                        new
                        {
                            text = "You are a concise translation assistant. Always return strict JSON. Keep translated text in the requested target language. Keep glossary words in the source language. Keep glossary meanings in the target language. Write nuance explanations in Korean.\n\n" + prompt
                        }
                    }
                }
            }
        };

        var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GetApiKey(settings)}";
        using var response = await client.PostAsync(
            url,
            new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"),
            cancellationToken);

        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();

        using var document = JsonDocument.Parse(body);
        return document.RootElement
            .GetProperty("candidates")[0]
            .GetProperty("content")
            .GetProperty("parts")[0]
            .GetProperty("text")
            .GetString() ?? throw new InvalidOperationException("Gemini returned no content.");
    }

    private async Task<string> RequestOpenAiImageAsync(AppSettings settings, string prompt, byte[] imageBytes, CancellationToken cancellationToken)
    {
        using var client = CreateClient(settings);
        var payload = new
        {
            model = "gpt-4o-mini",
            messages = new object[]
            {
                new
                {
                    role = "system",
                    content = "You are a concise OCR and translation assistant. Always return strict JSON."
                },
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "text", text = prompt },
                        new
                        {
                            type = "image_url",
                            image_url = new { url = $"data:image/png;base64,{Convert.ToBase64String(imageBytes)}" }
                        }
                    }
                }
            }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", GetApiKey(settings));
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using var response = await client.SendAsync(request, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();

        using var document = JsonDocument.Parse(body);
        return document.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString() ?? throw new InvalidOperationException("OpenAI returned no content for image translation.");
    }

    private async Task<string> RequestGeminiImageAsync(AppSettings settings, string prompt, byte[] imageBytes, CancellationToken cancellationToken)
    {
        using var client = CreateClient(settings);
        var payload = new
        {
            contents = new[]
            {
                new
                {
                    parts = new object[]
                    {
                        new { text = prompt },
                        new
                        {
                            inline_data = new
                            {
                                mime_type = "image/png",
                                data = Convert.ToBase64String(imageBytes)
                            }
                        }
                    }
                }
            }
        };

        var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GetApiKey(settings)}";
        using var response = await client.PostAsync(
            url,
            new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json"),
            cancellationToken);

        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();

        using var document = JsonDocument.Parse(body);
        return document.RootElement
            .GetProperty("candidates")[0]
            .GetProperty("content")
            .GetProperty("parts")[0]
            .GetProperty("text")
            .GetString() ?? throw new InvalidOperationException("Gemini returned no content for image translation.");
    }

    private static HttpClient CreateClient(AppSettings settings)
    {
        return new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(SettingsService.ClampTimeout(settings.TranslationTimeoutSeconds))
        };
    }

    private static string GetApiKey(AppSettings settings) => settings.SelectedProvider switch
    {
        AiProvider.Gemini => settings.GeminiApiKey,
        _ => settings.OpenAiApiKey
    };

    private static string BuildPrompt(string text, TranslationDirection direction)
    {
        return direction switch
        {
            TranslationDirection.KoreanToVietnamese => BuildDirectionalPrompt(text, "Korean", "Vietnamese"),
            TranslationDirection.EnglishToVietnamese => BuildDirectionalPrompt(text, "English", "Vietnamese"),
            TranslationDirection.VietnameseToKorean => BuildDirectionalPrompt(text, "Vietnamese", "Korean"),
            _ =>
                """
                Detect whether the following sentence is Korean or English, then translate it into Vietnamese.
                Return valid JSON only.

                Rules:
                - JSON shape:
                  {
                    "options": [
                      { "text": "...", "nuance": "..." },
                      { "text": "...", "nuance": "..." },
                      { "text": "...", "nuance": "..." }
                    ],
                    "glossary": [
                      { "word": "...", "meaning": "..." }
                    ]
                  }
                - Return exactly 3 translation options.
                - If the source text is Korean, treat it as Korean.
                - If the source text is English, treat it as English.
                - "text" must be written only in Vietnamese.
                - "nuance" must be brief, written only in Korean, and explain the tone difference.
                - "glossary" should contain 3 to 15 useful words or short phrases from the detected source sentence.
                - Each glossary "word" must stay in the original detected source language.
                - Each glossary "meaning" must be the matching word or short phrase in Vietnamese.
                - Do not add markdown fences or commentary.

                Sentence:
                """ + Environment.NewLine + text
        };
    }

    private static string BuildImagePrompt(TranslationDirection direction)
    {
        var target = direction switch
        {
            TranslationDirection.KoreanToVietnamese => ("Korean", "Vietnamese"),
            TranslationDirection.EnglishToVietnamese => ("English", "Vietnamese"),
            TranslationDirection.VietnameseToKorean => ("Vietnamese", "Korean"),
            _ => ("auto-detect Korean or English", "Vietnamese")
        };

        return $@"
Read the text from this image and translate it.
Return valid JSON only.

Rules:
- JSON shape:
  {{
    ""detectedText"": ""..."",
    ""options"": [
      {{ ""text"": ""..."", ""nuance"": ""..."" }},
      {{ ""text"": ""..."", ""nuance"": ""..."" }},
      {{ ""text"": ""..."", ""nuance"": ""..."" }}
    ],
    ""glossary"": [
      {{ ""word"": ""..."", ""meaning"": ""..."" }}
    ]
  }}
- ""detectedText"" must contain the extracted source text from the image.
- If direction is auto, detect whether the extracted text is Korean or English.
- Otherwise translate from {target.Item1} to {target.Item2}.
- Return exactly 3 translation options.
- ""nuance"" must be brief and written in Korean.
- ""glossary"" should contain 3 to 15 useful words or phrases from detectedText.
- Do not add markdown fences or commentary.
".Trim();
    }

    private static string BuildDirectionalPrompt(string text, string sourceLanguage, string targetLanguage) =>
        $@"
Translate the following sentence from {sourceLanguage} into {targetLanguage}.
Return valid JSON only.

Rules:
- JSON shape:
  {{
    ""options"": [
      {{ ""text"": ""..."", ""nuance"": ""..."" }},
      {{ ""text"": ""..."", ""nuance"": ""..."" }},
      {{ ""text"": ""..."", ""nuance"": ""..."" }}
    ],
    ""glossary"": [
      {{ ""word"": ""..."", ""meaning"": ""..."" }}
    ]
  }}
- Return exactly 3 translation options.
- ""text"" must be written only in {targetLanguage}.
- ""nuance"" must be brief, written only in Korean, and explain the tone difference.
- ""glossary"" should contain 3 to 15 useful words or short phrases from the original {sourceLanguage} sentence.
- Each glossary ""word"" must stay in the original {sourceLanguage}.
- Each glossary ""meaning"" must be the matching word or short phrase in {targetLanguage}.
- Format the glossary like a mini vocabulary list: source word -> target-language equivalent.
- Do not add markdown fences or commentary.

Sentence:
{text}
".Trim();
}
