using System.Globalization;
using md2visio.Localization;

namespace md2visio.Tests.Localization;

[CollectionDefinition("Culture-sensitive tests", DisableParallelization = true)]
public sealed class CultureSensitiveCollection;

[Collection("Culture-sensitive tests")]
public sealed class CoreStringsTests
{
    [Theory]
    [InlineData("en", "Conversion complete!")]
    [InlineData("zh-CN", "转换完成!")]
    public void Get_UsesRequestedUiCulture(string cultureName, string expected)
    {
        var originalCulture = CultureInfo.CurrentCulture;
        var originalUiCulture = CultureInfo.CurrentUICulture;

        try
        {
            var culture = CultureInfo.GetCultureInfo(cultureName);
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;

            Assert.Equal(expected, CoreStrings.Get("ConversionComplete"));
            Assert.DoesNotContain("[", CoreStrings.Format("GeneratedFiles", 2));
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUiCulture;
        }
    }
}
