using md2visio.mermaid.cmn;

namespace md2visio.mermaid.packet
{
    internal class PaSttChar : SynState
    {
        public override SynState NextState()
        {
            string? next = Ctx.Peek();
            if (next == null) return EndOfFile;

            if (next == "%") return Forward<SttPercent>();
            if (next == ":") return Forward<PacSttTuple>();
            if (next == "\n") return Forward<SttFinishFlag>();
            if (next == "`") return Forward<SttMermaidClose>();

            return Take().Forward<PaSttChar>();
        }
    }
}
