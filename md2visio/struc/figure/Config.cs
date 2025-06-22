namespace md2visio.struc.figure
{
    internal class Config: IConfig
    {
        ConfigDefaults configDefaults;
        MmdFrontMatter userFrontMatter = new();
        MmdJsonObj userDirective = new();

        public MmdFrontMatter UserFrontMatter { get => userFrontMatter; set => userFrontMatter = value; }
        public MmdJsonObj UserDirective { get => userDirective; set => userDirective = value; }

        public Config(Figure figure)
        {
            configDefaults = new(figure);
        }
        public MmdFrontMatter LoadUserDirective(string directive)
        {
            try
            {
                userDirective.Load(directive);
                return userFrontMatter.UpdateWith(UserDirective);
            }
            catch (ArgumentException)
            {
                return userFrontMatter;
            }
        }

        public MmdFrontMatter LoadUserDirectiveFromComment(string comment)
        {
            if (!comment.TrimEnd(' ').EndsWith("%%")) return Empty.Get<MmdFrontMatter>();

            char[] trims = ['%', '\t', '\n', '\r', '\f', ' '];
            string setting = comment.TrimStart(trims).TrimEnd(trims);
            return LoadUserDirective(setting);
        }
        public MmdFrontMatter LoadUserFrontMatter(string frontMatter)
        {
            return UserFrontMatter.LoadYaml(frontMatter);
        }

        public bool GetDouble(string keyPath, out double d)
        {
            return GetUserDouble(keyPath, out d)
                || UpdateDefaults().GetDouble(keyPath, out d);
        }
        public bool GetInt(string keyPath, out int i)
        {
            return GetUserInt(keyPath, out i)
                || UpdateDefaults().GetInt(keyPath, out i);
        }
        public bool GetString(string keyPath, out string s)
        {
            return GetUserString(keyPath, out s) 
                || UpdateDefaults().GetString(keyPath, out s);
        }

        public List<string> GetStringList(string prefix, int maxCount = 13)
        {
            List<string> serial = new List<string>();
            for (int i = 0; i < maxCount; ++i)
            {
                if (GetString($"{prefix}{i}", out string s))
                    serial.Add(s);
            }
            return serial;
        }

        bool GetUserString(string keyPath, out string s)
        {
            if (userDirective.GetString(keyPath, out s)) return true;
            return userFrontMatter.GetString(keyPath, out s);
        }

        bool GetUserInt(string keyPath, out int i)
        {
            if (userDirective.GetInt(keyPath, out i)) return true;
            return userFrontMatter.GetInt(keyPath, out i);
        }

        bool GetUserDouble(string keyPath, out double d)
        {
            if (userDirective.GetDouble(keyPath, out d)) return true;
            return userFrontMatter.GetDouble(keyPath, out d);
        }

        ConfigDefaults UpdateDefaults()
        {
            configDefaults.Theme = GetUserString("config.theme", out string theme) ? theme : "default";
            configDefaults.DarkMode = GetUserString("config.themeVariables.darkMode", out string darkMode) 
                && darkMode == "true";    

            return configDefaults;
        }
    }
}
