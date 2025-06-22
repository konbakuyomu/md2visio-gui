using md2visio.mermaid.cmn;
using md2visio.mermaid.graph;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph.@internal
{
    internal class GSttWord : SynState
    {
        public override SynState NextState()
        {
            // keyword or text
            if (string.IsNullOrWhiteSpace(Buffer)) throw new SynException("syntax error", Ctx);

            if (GSttKeyword.IsKeyword(Buffer)) return Forward<GSttKeyword>();
            else return Forward<GSttText>();
        }
    }
}
