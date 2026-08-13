using System.Globalization;

namespace md2visio.GUI.Localization;

internal static class CultureSettings
{
    private const string LanguageEnvironmentVariable = "MD2VISIO_LANGUAGE";
    private const string SettingsFileName = "language.txt";
    private static readonly string[] SupportedCultures = ["en", "zh-CN"];

    public static string CurrentCultureName =>
        CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase)
            ? "zh-CN"
            : "en";

    public static void ApplySavedCulture()
    {
        var requestedCulture = Environment.GetEnvironmentVariable(LanguageEnvironmentVariable);
        if (string.IsNullOrWhiteSpace(requestedCulture))
        {
            try
            {
                var settingsPath = GetSettingsPath();
                if (File.Exists(settingsPath))
                    requestedCulture = File.ReadAllText(settingsPath).Trim();
            }
            catch
            {
                // A read-only profile should not prevent the application from starting.
            }
        }

        SetCulture(Normalize(requestedCulture ?? CultureInfo.CurrentUICulture.Name));
    }

    public static void SaveAndApply(string cultureName)
    {
        cultureName = Normalize(cultureName);
        SetCulture(cultureName);

        try
        {
            var settingsPath = GetSettingsPath();
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
            File.WriteAllText(settingsPath, cultureName);
        }
        catch
        {
            // The selected language still applies for this session.
        }
    }

    private static void SetCulture(string cultureName)
    {
        var culture = CultureInfo.GetCultureInfo(cultureName);
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
    }

    private static string Normalize(string cultureName) =>
        SupportedCultures.Contains(cultureName, StringComparer.OrdinalIgnoreCase) ||
        cultureName.StartsWith("zh", StringComparison.OrdinalIgnoreCase)
            ? cultureName.StartsWith("zh", StringComparison.OrdinalIgnoreCase) ? "zh-CN" : "en"
            : "en";

    private static string GetSettingsPath() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "md2visio",
        SettingsFileName);
}
