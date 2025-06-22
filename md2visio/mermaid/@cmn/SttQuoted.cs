using System.Text;

namespace md2visio.mermaid.cmn
{
    internal abstract class SttQuoted : SynState
    {
        public static bool TestQuoted(SynContext ctx)
        {
            return ctx.Test(@"^(?s)""(?<quote>[^""]*)""");
        }

        public static string TrimQuote(string quoted)
        {
            StringBuilder sb = new StringBuilder(quoted.Trim());
            if (sb[0] == '"') sb.Remove(0, 1);
            if (sb.Length > 0 && sb[sb.Length - 1] == '"') sb.Remove(sb.Length - 1, 1);

            return sb.ToString();            
        }
    }
}
