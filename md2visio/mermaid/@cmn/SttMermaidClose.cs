using System.Text.RegularExpressions;

namespace md2visio.mermaid.cmn
{
    internal class SttMermaidClose : SynState
    {
        public override SynState NextState()
        {
            if (!IsMermaidClose(Ctx)) throw new SynException($"expected mermaid close", Ctx);

            string closeTag = Ctx.TestGroups["bquote"].Value;
            return Save(closeTag).Slide(Ctx.TestGroups[0].Value.Length).Forward<SttMermaidStart>();
        }

        public static bool IsMermaidClose(SynContext ctx)
        {
            (bool success, SynState mermaidStart) = ctx.FindContainerType("SttMermaidStart");
            if (!success) return false;

            int L = mermaidStart.Fragment.Length;
            return ctx.Test($@"^\s*(?<bquote>`{{{L},}})\s*(?=\n|$)");
        }
    }
}
