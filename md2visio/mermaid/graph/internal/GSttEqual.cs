using md2visio.mermaid.cmn;

namespace md2visio.mermaid.graph.@internal
{
    internal class GSttEqual : SynState
    {
        public override SynState NextState()
        {
            // left
            if (Buffer.Length > 0)
            {
                if (GSttLinkStart.IsLinkStart(Ctx))
                    return Restore(Buffer.Length).ClearBuffer().Forward<GSttLinkStart>();
                if (GSttNoLabelLink.IsNoLabelLink(Ctx))
                    return Restore(Buffer.Length).ClearBuffer().Forward<GSttNoLabelLink>();

                throw new SynException("unexpected '='", Ctx);
            }

            // right
            if (GSttLinkStart.IsLinkStart(Ctx)) return Forward<GSttLinkStart>();
            else if (GSttNoLabelLink.IsNoLabelLink(Ctx)) return Forward<GSttNoLabelLink>();

            throw new SynException("unexpected '='", Ctx);
        }
    }
}
