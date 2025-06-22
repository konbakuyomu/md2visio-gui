using System.Text;

namespace md2visio.mermaid.cmn
{
    internal class CompoDict
    {
        public static CompoDict Empty = new CompoDict();
        string entire = "entire";
        List<CompoValue> valueList = new List<CompoValue>(); // key the order of added values ([0]: Entire)
        Dictionary<string, CompoValue> dict = new Dictionary<string, CompoValue>();

        public string Entire { get { return valueList[0].Value; } set { valueList[0].Value = value; } }

        public CompoDict()
        {
            Add(entire, string.Empty);
        }

        public void Add(string key, string value)
        {
            if (dict.ContainsKey(key)) dict[key].Value = value;
            else
            {
                CompoValue pv = new CompoValue(key, value);
                valueList.Add(pv);
                dict.Add(key, pv);
            }
        }

        public void Add(string value)
        {
            Add($"{valueList.Count}", value);
        }

        public bool ContainsKey(object key)
        {
            return dict.ContainsKey($"{key}");
        }

        public string Get(string key)
        {
            return dict[key].Value;
        }

        public string Get(object key)
        {
            return Get($"{key}");
        }

        public string? this[object key]
        {
            get
            {
                if (!ContainsKey(key)) return null;
                return Get($"{key}");
            }
        }

        public List<CompoValue> Values()
        {
            return valueList;
        }

        public string MakeWhole()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var item in valueList)
            {
                if (item.Key == entire) continue;
                sb.Append(item.Value).Append(' ');
            }
            return sb.ToString();
        }
    }
}
