using md2visio.mermaid.cmn;
using md2visio.mermaid.graph.@internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace md2visio.mermaid.graph
{
    internal class GSttAmp : SynState
    {
        public override SynState NextState()
        {
            string? pre = Peek(-1), next2 = Peek(2);
            if (!string.IsNullOrWhiteSpace(pre)) return Take().Forward<GSttChar>();
            if ("&" != next2?.Trim()) return Take().Forward<GSttChar>();

            if (Ctx.LastNonFinishState() is not GSttText) throw new SynException("syntax error", Ctx);

            return Save("&").Slide(2).Forward<GSttChar>();
        }
    }
}
