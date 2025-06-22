using md2visio.mermaid.cmn;
using md2visio.mermaid.graph.@internal;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph
{
    internal class GSttPipedLinkText : SynState
    {
        public override SynState NextState()
        {
            if (!IsPipedLinkText(Ctx)) throw new SynException("expected piped link text", Ctx);

            Group textGroup = Ctx.TestGroups["text"];
            string text = textGroup.Value.Trim();
            return Save(text).Slide(Ctx.TestGroups[0].Length).Forward<GSttChar>();
        }

        public static bool IsPipedLinkText(SynContext ctx)
        {
            if (!ctx.Test(@"(?s)^\s*\|")) return false;
            if (!ctx.Test(@"(?s)^\s*\|(?<text>.+?)\|")) return false;
            return true;
        }
    }
}
