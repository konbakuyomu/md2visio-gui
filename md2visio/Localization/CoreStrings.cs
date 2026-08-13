using System.Globalization;
using System.Resources;

namespace md2visio.Localization;

public static class CoreStrings
{
    private static readonly ResourceManager Resources =
        new("md2visio.Resources.Strings", typeof(CoreStrings).Assembly);

    public static string Get(string key) =>
        Resources.GetString(key, CultureInfo.CurrentUICulture) ?? $"[{key}]";

    public static string Format(string key, params object?[] args) =>
        string.Format(CultureInfo.CurrentCulture, Get(key), args);
}
