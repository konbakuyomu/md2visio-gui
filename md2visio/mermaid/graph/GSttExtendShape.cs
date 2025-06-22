using md2visio.mermaid.cmn;
using md2visio.mermaid.graph.@internal;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph
{
    internal class GSttExtendShape : SynState
    {
        public override SynState NextState()
        {
            string? pre = Peek(-1);
            if (string.IsNullOrWhiteSpace(pre)) throw new SynException("syntax error", Ctx);
            if (!Expect(@"@\{")) throw new SynException("expected @{...}", Ctx);

            Restore();
            if (!GSttPaired.IsPaired("{", Ctx)) throw new SynException("expected @{...}", Ctx);

            Group pair = Ctx.TestGroups[0];
            return Save(pair.Value).Slide(pair.Length).Forward<GSttChar>();
        }
    }
}
