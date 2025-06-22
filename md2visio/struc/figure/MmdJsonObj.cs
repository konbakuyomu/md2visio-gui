using md2visio.mermaid.cmn;
using System.Text;

namespace md2visio.struc.figure
{
    internal class MmdJsonObj : ValueAccessor
    {
        Dictionary<string, object> data = new Dictionary<string, object>(); // string -> string/MmdJsonObj/MmdJsonArray
        int index = 0;

        public int Index { get { return index; } }
        public Dictionary<string, object> Data { get { return data; } }
        public MmdJsonObj() { }

        public MmdJsonObj(string text)
        {
            Load(new StringBuilder(text));
        }

        public MmdJsonObj(StringBuilder textBuilder, int index)
        {
            this.index = index;
            Load(textBuilder);
        }

        public MmdJsonObj Load(string text)
        {
            data.Clear();
            return Load(new StringBuilder(text));
        }

        public override T? GetValue<T>(string keyPath) where T : class
        {
            if (!keyPath.Contains('.'))
            {
                if (this[keyPath] is T) return data[keyPath] as T;
            }
            else
            {
                string[] path = keyPath.Split('.');
                int count = path.Length;
                object? result = this;
                foreach (string pathItem in path)
                {
                    if (result == null) break;

                    --count;
                    if (result is MmdJsonArray) result = (result as MmdJsonArray)?[pathItem];
                    else if (result is MmdJsonObj) result = (result as MmdJsonObj)?[pathItem];
                }
                if (count == 0 && result is T) return result as T;
            }
            return null;
        }

        public override void SetValue(string keyPath, object val)
        {
            string[] keyList = keyPath.Split(".");
            ValueAccessor? result = this;
            for (int i = 0; i < keyList.Length; ++i)
            {
                string key = keyList[i];
                if (result == null) break;

                if (i == keyList.Length - 1)
                {
                    if (result is MmdJsonObj) ((MmdJsonObj)result)[key] = val;
                    else if (result is MmdJsonArray) ((MmdJsonArray)result)[key] = val;
                }
                else
                {
                    ValueAccessor? next = result.GetValue<ValueAccessor>(key);
                    if (next == null)
                    {
                        MmdJsonObj newJson = new MmdJsonObj();
                        result.SetValue(key, newJson);
                        result = newJson;
                    }
                    else result = next;
                }
            }
        }

        public object? this[string key]
        {
            get { return data.TryGetValue(key, out object? value) ? value : null; }
            set
            {
                data[key] = value ?? string.Empty;
            }
        }

        public bool HasKey(string key)
        {
            return data.ContainsKey(key);
        }

        public MmdJsonObj UpdateWith(MmdJsonObj json)
        {
            return UpdateWith(json, new StringBuilder());
        }

        MmdJsonObj UpdateWith(MmdJsonObj json, StringBuilder path)
        {
            if (json == null) return this;

            foreach (string key in json.Data.Keys)
            {
                AppendKey(path, key);

                object? val = json[key];
                if (val is MmdJsonObj) UpdateWith((MmdJsonObj)val, new StringBuilder(path.ToString()));
                else SetValue(path.ToString(), val ?? string.Empty);

                UnappendKey(path);
            }

            return this;
        }

        MmdJsonObj Load(StringBuilder textBuilder)
        {
            StringBuilder keyBuilder = new StringBuilder();
            StringBuilder valueBuilder = new StringBuilder();
            bool withInQuote = false;
            bool withInSQuote = false;
            bool bAppendKey = true;
            for (; index < textBuilder.Length; ++index)
            {
                char c = textBuilder[index];
                if (c == '"') withInQuote = !withInQuote;
                else if (c == '\'') withInSQuote = !withInSQuote;

                if (withInQuote || withInSQuote)
                {
                    Append(keyBuilder, valueBuilder, c, bAppendKey);
                    continue;
                }

                if (c == ' ') continue;
                else if (c == ':')
                {
                    Assert($"syntax error near '{valueBuilder}'", TrimSpaceAndQuote(valueBuilder).Length == 0);
                    Assert("JSON key can't be empty", TrimSpaceAndQuote(keyBuilder).Length > 0);
                    bAppendKey = !bAppendKey;
                    continue;
                }
                else if (c == ',')
                {
                    bAppendKey = !bAppendKey;
                    AddString(keyBuilder, valueBuilder);
                    continue;
                }
                else if (c == '{')
                {
                    if (TrimSpaceAndQuote(keyBuilder).Length > 0)
                    {
                        MmdJsonObj obj = new MmdJsonObj(textBuilder, index);
                        AddJsonObj(keyBuilder, obj);
                        index = obj.Index;
                    }
                    continue;
                }
                else if (c == '}')
                {
                    AddString(keyBuilder, valueBuilder);
                    return this;
                }
                else if (c == '[')
                {
                    MmdJsonArray arr = new MmdJsonArray(textBuilder, index);
                    AddJsonObj(keyBuilder, arr);
                    index = arr.Index;
                    continue;
                }
                else if (c == ']')
                {
                    AddString(keyBuilder, valueBuilder);
                    continue;
                }

                Append(keyBuilder, valueBuilder, c, bAppendKey);
            }
            AddString(keyBuilder, valueBuilder); // not closed by '}'

            return this;
        }

        void AddString(StringBuilder key, StringBuilder value)
        {
            string k = TrimSpaceAndQuote(key);
            if (k.Length == 0) return;

            string v = TrimSpaceAndQuote(value);
            Assert("JSON value can't be empty", !string.IsNullOrEmpty(v));
            data[k] = v;
            key.Clear();
            value.Clear();
        }

        void AddJsonObj(StringBuilder key, object v)
        {
            string k = TrimSpaceAndQuote(key);
            Assert("JSON key can't be empty", !string.IsNullOrEmpty(k));

            data[k] = v;
            key.Clear();
        }

        string TrimSpaceAndQuote(StringBuilder key)
        {
            char[] trims = ['"', '\''];
            return key.ToString().Trim().TrimStart(trims).TrimEnd(trims);
        }

        void Append(StringBuilder key, StringBuilder value, char c, bool appendKey)
        {
            if (appendKey) key.Append(c);
            else value.Append(c);
        }

        void Assert(string message, bool assert)
        {
            if (!assert)
                throw new ArgumentException(message);
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder("{");
            List<string> keys = [.. data.Keys];
            foreach (string key in keys)
            {
                string quote = data[key] is ValueAccessor ? string.Empty : "'";
                sb.Append($"'{key}'").Append(": ")
                    .Append($"{quote}{data[key]}{quote}")
                    .Append(key == keys.Last() ? string.Empty : ", ");
            }
            return $"{sb}}}";
        }
    }
}
