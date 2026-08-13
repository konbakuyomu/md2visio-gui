using md2visio.vsdx.@base;

namespace md2visio.Tests.VisioParsing;

public sealed class VisioValueParserTests
{
    [Theory]
    [InlineData("2.8222 mm.", 2.8222)]
    [InlineData("2,8222 mm", 2.8222)]
    [InlineData("9 pt.", 9)]
    [InlineData("-1.25 in.", -1.25)]
    [InlineData("1.2E-3 mm", 0.0012)]
    public void ParseResult_AcceptsLocalizedVisioUnitStrings(string value, double expected)
    {
        Assert.Equal(expected, VisioValueParser.ParseResult(value), precision: 8);
    }

    [Fact]
    public void ParseResult_RejectsMissingNumbers()
    {
        Assert.Throws<FormatException>(() => VisioValueParser.ParseResult("mm."));
    }
}
