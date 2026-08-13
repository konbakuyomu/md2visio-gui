namespace md2visio.GUI.Localization;

internal static class OutputDirectorySettings
{
    private const string SettingsFileName = "output-directory.txt";

    public static string Load()
    {
        var fallback = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

        try
        {
            var settingsPath = GetSettingsPath();
            if (!File.Exists(settingsPath)) return fallback;

            var savedDirectory = File.ReadAllText(settingsPath).Trim();
            return Directory.Exists(savedDirectory) ? savedDirectory : fallback;
        }
        catch
        {
            return fallback;
        }
    }

    public static void Save(string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory)) return;

        try
        {
            var fullPath = Path.GetFullPath(directory.Trim());
            var settingsPath = GetSettingsPath();
            Directory.CreateDirectory(Path.GetDirectoryName(settingsPath)!);
            File.WriteAllText(settingsPath, fullPath);
        }
        catch
        {
            // A read-only profile should not interrupt normal application use.
        }
    }

    private static string GetSettingsPath() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "md2visio",
        SettingsFileName);
}
