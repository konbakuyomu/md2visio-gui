namespace md2visio.struc.figure
{
    internal interface IConfig
    {
        bool GetDouble(string keyPath, out double d);
        bool GetInt(string keyPath, out int i);
        bool GetString(string keyPath, out string val);
    }
}
