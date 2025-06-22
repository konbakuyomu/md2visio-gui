using md2visio.mermaid.cmn;

namespace md2visio.mermaid.graph.@internal
{
    internal class GSttBackQuote : SynState
    {
        public override SynState NextState()
        {
            if (SttMermaidClose.IsMermaidClose(Ctx)) return Forward<SttMermaidClose>();

            return Take().Forward<GSttChar>();
        }
    }
}
