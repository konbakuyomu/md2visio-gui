using md2visio.mermaid.cmn;

namespace md2visio.struc.figure
{
    internal class ConfigDefaults: IConfig
    {
        static readonly string defaultDir = "./default";
        static readonly string themeDir = $"{defaultDir}/theme";        
        static readonly Dictionary<string, MmdFrontMatter> defConfig = new(); // ./default/?.yaml
        static readonly Dictionary<string, MmdFrontMatter> themeVars = new(); // ./default/theme/?.yaml
        static readonly string commonCfg = "default"; // ./default/default.yaml

        public Figure Figure { get; }
        public string Theme { get; set; } = "default";
        public bool DarkMode { get; set; } = false;

        public ConfigDefaults(Figure figure)
        {
            Figure = figure;
            LoadFigureConfig(figure);
            LoadCommonConfig();            
        }

        public bool GetDouble(string keyPath, out double d)
        {
            return LoadThemeVars(Theme, DarkMode).GetDouble(keyPath, out d)
                || LoadFigureConfig(Figure).GetDouble(keyPath, out d)
                || LoadCommonConfig().GetDouble(keyPath, out d);
        }
        public bool GetInt(string keyPath, out int i)
        {
            return LoadThemeVars(Theme, DarkMode).GetInt(keyPath, out i)
                || LoadFigureConfig(Figure).GetInt(keyPath, out i)
                || LoadCommonConfig().GetInt(keyPath, out i);
        }
        public bool GetString(string keyPath, out string s)
        {
            return LoadThemeVars(Theme, DarkMode).GetString(keyPath, out s)
                || LoadFigureConfig(Figure).GetString(keyPath, out s)
                || LoadCommonConfig().GetString(keyPath, out s);
        }

        MmdFrontMatter LoadFigureConfig(Figure figure)
        {
            if (!TypeMap.ConfigMap.TryGetValue(figure.GetType().Name, out string? fName))
                return Empty.Get<MmdFrontMatter>();

            if (defConfig.TryGetValue(fName, out MmdFrontMatter? fm))
                return fm ?? Empty.Get<MmdFrontMatter>();

            fm = MmdFrontMatter.FromFile($"{defaultDir}/{fName}.yaml");
            defConfig.Add(fName, fm);
            return fm;
        }

        MmdFrontMatter LoadCommonConfig()
        {
            if (defConfig.TryGetValue(commonCfg, out MmdFrontMatter? fm))
                return fm;

            fm = MmdFrontMatter.FromFile($"{defaultDir}/{commonCfg}.yaml");
            defConfig.Add(commonCfg, fm);
            return fm;
        }
        MmdFrontMatter LoadThemeVars(string theme2load, bool darkMode = false)
        {
            DarkMode = darkMode;
            Theme = theme2load.ToLower();
            Theme = string.Format("{0}{1}", Theme, Theme == "base" && darkMode ? "-darkMode" : "");
            if (themeVars.TryGetValue(Theme, out MmdFrontMatter? fm))
                return fm ?? Empty.Get<MmdFrontMatter>();

            fm = MmdFrontMatter.FromFile($"{themeDir}/{Theme}.yaml");
            themeVars.Add(Theme, fm);

            return fm;
        }
    }
}
