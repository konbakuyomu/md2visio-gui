
namespace md2visio.vsdx.@base
{
    public static class VDrawerAssert
    {
        public static void Assert<T>(this T obj, bool assert, string message) where T : class
        {
            if(!assert) throw new VDrawerException(message);
        }
    }
}
