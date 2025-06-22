using md2visio.mermaid.cmn;

namespace md2visio.mermaid.xy
{
    internal class XySttChar : SynState
    {
        public override SynState NextState()
        {
            string? next = Ctx.Peek();
            if (next == null) return EndOfFile;

            if (next == "%") return Forward<SttPercent>();
            if (next == " ") return Forward<XySttWord>();
            if (next == "\t") return Forward<XySttWord>();
            if (next == "\n") return Forward<SttFinishFlag>();
            if (next == "`") return Forward<SttMermaidClose>();

            return Take().Forward<XySttChar>();
        }
    }
}
