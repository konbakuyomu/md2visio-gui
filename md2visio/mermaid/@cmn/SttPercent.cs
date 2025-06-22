using md2visio.mermaid.graph.@internal;

namespace md2visio.mermaid.cmn
{
    internal class SttPercent : SynState
    {
        public override SynState NextState()
        {
            if (string.IsNullOrWhiteSpace(Buffer))
            {
                if (Expect("%%")) return Restore(2).Forward<SttComment>();
            }

            return Take().Forward<SttCtxChar>();
        }
    }
}
