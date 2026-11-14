using PortalGNLAB.Web.Services;
using Xunit;

namespace PortalGNLAB.Db.Tests;

public sealed class TeamMemoTranslationServiceTests
{
    private const string Earlier = "gpt-luna · low";
    private const string Now = "gpt-luna · low · " + TeamMemoTranslationService.PromptVersion;

    [Fact]
    public void KeptAsWritten_AsksAgainForAVietnameseProcessNameGivenAsItsEnglish()
    {
        // as cached for the Process table: the English of "Xếp UPPER" was "Xếp UPPER"
        Assert.True(TeamMemoTranslationService.KeptAsWritten("Xếp UPPER", "en", "Xếp UPPER", Earlier));
        Assert.True(TeamMemoTranslationService.KeptAsWritten("Kiểm tra AOI lần 1", "ko", "Kiểm tra AOI lần 1", Earlier));
        // a real translation, the Vietnamese as written, a code, or one asked again with this prompt are kept
        Assert.False(TeamMemoTranslationService.KeptAsWritten("Xếp SPK", "en", "Arrange SPK", Earlier));
        Assert.False(TeamMemoTranslationService.KeptAsWritten("Xếp UPPER", "vi", "Xếp UPPER", Earlier));
        Assert.False(TeamMemoTranslationService.KeptAsWritten("DRY", "en", "DRY", Earlier));
        Assert.False(TeamMemoTranslationService.KeptAsWritten("SPK 조립", "en", "SPK 조립", Earlier));
        Assert.False(TeamMemoTranslationService.KeptAsWritten("Xếp UPPER", "en", "Xếp UPPER", Now));
    }

    [Fact]
    public void Instructions_TranslateVietnameseAndKeepPartNamesInCapitals()
    {
        Assert.Contains("Latin letters but is not English", TeamMemoTranslationService.Instructions);
        Assert.Contains("\"Xếp UPPER\" is \"Stack UPPER\"", TeamMemoTranslationService.Instructions);
        Assert.Contains("UPPER / LOWER name a part", TeamMemoTranslationService.Instructions);
        Assert.DoesNotContain("process names in Latin letters", TeamMemoTranslationService.Instructions);
    }
}
