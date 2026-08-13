using System.Globalization;
using System.Resources;

namespace md2visio.GUI.Localization;

internal static class UiStrings
{
    private static readonly ResourceManager Resources =
        new("md2visio.GUI.Resources.Strings", typeof(UiStrings).Assembly);

    public static string Get(string key) =>
        Resources.GetString(key, CultureInfo.CurrentUICulture) ?? $"[{key}]";

    public static string Format(string key, params object?[] args) =>
        string.Format(CultureInfo.CurrentCulture, Get(key), args);
}
