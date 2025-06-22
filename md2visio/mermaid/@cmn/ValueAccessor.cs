using System.Text;

namespace md2visio.mermaid.cmn
{
    internal abstract class ValueAccessor
    {
        abstract public T? GetValue<T>(string keyPath) where T : class;
        abstract public void SetValue(string key, object value);

        virtual public bool GetString(string keyPath, out string s)
        {
            s = GetValue<string>(keyPath) ?? string.Empty;
            return s != string.Empty;
        }

        virtual public bool GetInt(string keyPath, out int i)
        {
            string? val = GetValue<string>(keyPath);
            bool success = int.TryParse(val, out i);
            return success;
        }

        virtual public bool GetDouble(string keyPath, out double d)
        {
            string? val = GetValue<string>(keyPath);
            bool success = double.TryParse(val, out d);
            return success;
        }

        protected void AppendKey(StringBuilder path, string key)
        {
            if (path.Length > 0) path.Append(".");
            path.Append(key);
        }

        protected void UnappendKey(StringBuilder path)
        {
            int dotIndex = path.Length - 1;
            for (; dotIndex >= 0; --dotIndex)
            {
                if (path[dotIndex] == '.') break;                
            }
            if(dotIndex >= 0) path.Remove(dotIndex, path.Length - dotIndex);
            else path.Clear();
        }
    }
}
