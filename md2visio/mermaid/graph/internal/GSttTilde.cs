using md2visio.mermaid.cmn;

namespace md2visio.mermaid.graph.@internal
{
    internal class GSttTilde : SynState
    {
        public override SynState NextState()
        {
            if (GSttNoLabelLink.IsEmptyLink(Ctx)) return Forward<GSttNoLabelLink>();

            throw new SynException("unexpected '~'", Ctx);
        }
    }
}
