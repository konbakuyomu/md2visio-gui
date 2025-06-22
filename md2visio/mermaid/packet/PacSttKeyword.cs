using md2visio.mermaid.cmn;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.packet
{
    internal class PacSttKeyword : SynState
    {
        public override SynState NextState()
        {
            Save(Buffer).ClearBuffer();
            return Forward<PaSttChar>();
        }

        public static bool IsKeyword(SynContext ctx)
        {
            return Regex.IsMatch(ctx.Cache.ToString(), "^(packet-data|packet)$");
        }
    }
}
