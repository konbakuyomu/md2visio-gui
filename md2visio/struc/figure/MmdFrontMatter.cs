using md2visio.mermaid.cmn;
using System.Text;
using YamlDotNet.Serialization;

namespace md2visio.struc.figure
{
    internal class MmdFrontMatter : ValueAccessor
    {
        object? yamlObj;

        public MmdFrontMatter() {
            LoadYaml("{}");
        }

        public MmdFrontMatter(string yaml)
        {
            LoadYaml(yaml);
        }

        public MmdFrontMatter(object? yamlObj)
        {
            this.yamlObj = yamlObj;
        }

        public static MmdFrontMatter FromYaml(string yaml)
        {
            return new MmdFrontMatter(yaml);
        }

        public static MmdFrontMatter FromFile(string filePath)
        {
            return new MmdFrontMatter().LoadFile(filePath);
        }

        public MmdFrontMatter LoadYaml(string yaml)
        {
            yamlObj = new DeserializerBuilder().Build().Deserialize(yaml);
            return this;
        }

        public MmdFrontMatter LoadFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"config file '{filePath}' doesn't exist, use empty config instead");
                return new MmdFrontMatter();
            }

            StringBuilder yaml = new StringBuilder();
            foreach (var line in File.ReadAllLines(filePath))
                yaml.AppendLine(line);
            return LoadYaml(yaml.ToString());
        }

        public override T? GetValue<T>(string keyPath) where T : class
        {
            string[] items = keyPath.Split('.');
            int count = items.Length;
            MmdFrontMatter fm = this;
            object? result = yamlObj;
            foreach (string key in items)
            {
                --count;
                result = fm.GetItem(key);
                if (result == null) break;

                fm = new MmdFrontMatter(result);
            }
            if (count == 0 && result is T) return result as T;

            return null;
        }

        public override void SetValue(string keyPath, object value)
        {
            string[] keyList = keyPath.Split('.');
            MmdFrontMatter fm = this;
            for (int i = 0; i < keyList.Length; ++i)
            {
                string key = keyList[i];
                if (i == keyList.Length - 1) fm[key] = value;
                else
                {
                    object? result = fm.GetItem(key);
                    if (result == null)
                    {
                        MmdFrontMatter newFM = new("{}");
                        if (newFM.yamlObj == null) break;

                        fm.SetItem(key, newFM.yamlObj);
                        fm = newFM;
                    }
                    else fm = new MmdFrontMatter(result);
                }
            }
        }

        object? GetItem(string key)
        {
            if (yamlObj is IList<object>)
            {
                if (!int.TryParse(key, out int index)) return null;

                return GetItem(index);
            }
            else if (yamlObj is IDictionary<object, object>)
            {
                if (yamlObj is IDictionary<object, object> dict && dict.ContainsKey(key)) return dict[key];
            }
            return null;
        }

        public object? this[string key]
        {
            get => GetItem(key);
            set => SetItem(key, value);
        }

        public object? this[int index]
        {
            get => GetItem(index);
            set => SetItem(index, value);
        }

        void SetItem(string key, object? obj)
        {
            if (yamlObj is IList<object>)
            {
                if (int.TryParse(key, out int index))
                {
                    SetItem(index, obj);
                    return;
                }
            }
            else if (yamlObj is IDictionary<object, object>)
            {
                if (yamlObj is IDictionary<object, object> dict) dict[key] = obj ?? string.Empty;
            }
        }

        object? GetItem(int index)
        {
            if (yamlObj is not IList<object> list) return null;
            if (index < 0 || index >= list.Count) return null;

            return list[index];
        }

        void SetItem(int index, object? obj)
        {
            if (yamlObj is not IList<object> arr) return;

            int num2add = index - arr.Count + 1;
            for (int i = 0; i < num2add; i++)
                arr.Insert(arr.Count, string.Empty);

            arr[index] = obj ?? string.Empty;
        }

        public MmdFrontMatter UpdateWith(MmdJsonObj directive)
        {
            MmdJsonObj? init = directive.GetValue<MmdJsonObj>("init");
            if (init == null) return this;

            MmdJsonObj config = new("config: {}");
            config.SetValue("config", init);

            return UpdateWith(config, new StringBuilder());
        }

        MmdFrontMatter UpdateWith(MmdJsonObj json, StringBuilder path)
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

        public MmdFrontMatter UpdateWith(MmdFrontMatter another)
        {
            if (another.yamlObj == null) return this;

            return Join(another.yamlObj, new StringBuilder());
        }

        MmdFrontMatter Join(object yaml, StringBuilder path)
        {
            if (yaml is not IDictionary<object, object> dict) return this;

            foreach (object key in dict.Keys)
            {
                AppendKey(path, key.ToString() ?? string.Empty);
                if (dict == null) break;

                object? val = dict[key];
                if (val is IDictionary<object, object>) Join(val, new StringBuilder(path.ToString()));
                else SetValue(path.ToString(), val);

                UnappendKey(path);
            }

            return this;
        }

        public override string ToString()
        {
            if (yamlObj == null) return string.Empty;

            return ToString(yamlObj);
        }

        string ToString(object obj)
        {
            if (obj is IList<object>)
            {
                StringBuilder sb = new("[");
                IList<object> list = (IList<object>)obj;
                foreach (object item in list)
                {
                    sb.Append(ToString(item))
                        .Append(item == list.Last() ? string.Empty : ", ");
                }
                return $"{sb}]";
            }
            else if (obj is IDictionary<object, object>)
            {
                StringBuilder sb = new("{");
                IDictionary<object, object> dict = (IDictionary<object, object>)obj;
                List<object> keys = [.. dict.Keys];
                foreach (object key in keys)
                {
                    sb.Append(ToString(key)).Append(": ")
                        .Append(ToString(dict[key]))
                        .Append(key == keys.Last() ? string.Empty : ", ");
                }
                return $"{sb}}}";
            }

            return $"{Quote(obj)}{obj}{Quote(obj)}";
        }

        string Quote(object obj)
        {
            if (obj is IList<object> || obj is IDictionary<object, object>) return string.Empty;
            if (obj is string) return "'";
            return string.Empty;
        }

    }
}
