using md2visio.mermaid.cmn;
using System.Collections;
using System.Text;

namespace md2visio.struc.figure
{
    internal class MmdJsonArray : ValueAccessor, IEnumerable<object>
    {
        readonly List<object> list = [];
        int index = 0;
        public int Index { get { return index; } }
        public int Count {  get { return list.Count; } }

        public MmdJsonArray() { }

        public MmdJsonArray(string json)
        {
            Load(json);
        }

        public MmdJsonArray(StringBuilder textBuilder, int index)
        {
            this.index = index;
            Load(textBuilder);
        }

        public override T? GetValue<T>(string keyPath) where T : class
        {
            if (this[keyPath] is T) return this[keyPath] as T;
            return null;
        }

        public override void SetValue(string key, object value)
        {
            if (!int.TryParse(key, out int index)) return;

            int num2add = index - list.Count + 1;
            for (int i = 0; i < num2add; i++)
                list.Insert(list.Count, string.Empty);

            list[index] = value;
        }

        public object? this[string key]
        {
            get
            {
                if (!int.TryParse(key, out int v)) return null;
                return this[v];
            }
            set
            {
                if (int.TryParse(key, out int v)) this[v] = value;
            }
        }

        public object? this[int index]
        {
            get { return index >= 0 && index < list.Count ? list[index] : null; }
            set
            {
                int num = index - list.Count + 1;
                for (int i = 0; i < num; i++)
                    list.Insert(list.Count, string.Empty);
                list[index] = value ?? string.Empty;
            }
        }

        public MmdJsonArray Load(string json)
        {
            index = 0;
            return Load(new StringBuilder(json));
        }

        MmdJsonArray Load(StringBuilder textBuilder)
        {
            StringBuilder item = new();
            bool withInQuote = false;
            bool withInSQuote = false;
            for (; index < textBuilder.Length; ++index)
            {
                char c = textBuilder[index];
                if (c == '"') withInQuote = !withInQuote;
                else if (c == '\'') withInSQuote = !withInSQuote;

                if (withInQuote || withInSQuote)
                {
                    item.Append(c);
                    continue;
                }

                if (c == ',')
                {
                    AddString(item);
                    continue;
                }
                else if (c == '{')
                {
                    Assert($"syntax error near '{item}'", TrimSpaceAndQuote(item).Length == 0);

                    MmdJsonObj obj = new(textBuilder, index);
                    AddJsonObj(obj);
                    index = obj.Index;
                    continue;
                }
                else if (c == '[')
                {
                    Assert($"syntax error near '{item}'", TrimSpaceAndQuote(item).Length == 0);

                    if (list.Count == 0) { continue; }

                    MmdJsonArray arr = new(textBuilder, index + 1);
                    AddJsonObj(arr);
                    index = arr.Index;
                    continue;
                }
                else if (c == ']')
                {
                    AddString(item);
                    return this;
                }
                else if (c == ' ') continue;

                item.Append(c);
            }
            AddString(item); // not closed by ']'

            return this;
        }

        void AddString(StringBuilder item)
        {
            string v = TrimSpaceAndQuote(item);
            if (v.Length > 0) list.Add(v);
            item.Clear();
        }

        void AddJsonObj(object value)
        {
            list.Add(value);
        }

        string TrimSpaceAndQuote(StringBuilder stringBuilder)
        {
            char[] trims = ['"', '\''];
            return stringBuilder.ToString().Trim().TrimStart(trims).TrimEnd(trims);
        }

        void Assert(string message, bool test)
        {
            if (!test) throw new ArgumentException(message);
        }

        public override string ToString()
        {
            StringBuilder sb = new("[");
            foreach (var item in list)
            {
                string quote = item is ValueAccessor ? string.Empty : "'";
                sb.Append($"{quote}{item}{quote}")
                    .Append(item == list.Last() ? string.Empty : ", ");
            }
            return $"{sb}]";
        }

        public IEnumerator<object> GetEnumerator()
        {
            return new MmdJsonArrayEnumerator(list);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return new MmdJsonArrayEnumerator(list);
        }

        private class MmdJsonArrayEnumerator : IEnumerator<object>
        {
            private List<object> data;
            private int index = -1;

            public MmdJsonArrayEnumerator(List<object> data)
            {
                this.data = data;
            }

            public object Current => data[index];

            object IEnumerator.Current => Current;

            public void Dispose() { }

            public bool MoveNext()
            {
                index++;
                return index < data.Count;
            }

            public void Reset()
            {
                index = -1;
            }
        }
    }
}
