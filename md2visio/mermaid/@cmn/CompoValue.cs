namespace md2visio.mermaid.cmn
{
    internal class CompoValue
    {
        string key = string.Empty;
        string value = string.Empty;
        public string Key { get { return key; } set { key = value; } }
        public string Value { get { return value; } set { this.value = value; } }

        public CompoValue(string key, string value)
        {
            this.key = key;
            this.value = value;
        }

        public bool IsQuoted() { return key == "\""; }
        public bool IsPaired() { return key == "["; }
        public bool IsText() { return key == "T"; }
        public bool IsList() { return key == ","; }

    }
}
