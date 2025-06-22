using md2visio.mermaid.cmn;

namespace md2visio.mermaid.pie
{
    internal class PieSttChar : SynState
    {
        public override SynState NextState()
        {
            string? next = Ctx.Peek();
            if (next == null) return EndOfFile;

            if (next == "%")  return Forward<SttPercent>();
            if (next == ":")  return Forward<PieSttTuple>();
            if (next == " ")  return Forward<PieSttWord>();
            if (next == "\t") return Forward<PieSttWord>();
            if (next == "\n") return Forward<SttFinishFlag>();
            if (next == "`")  return Forward<SttMermaidClose>();

            return Take().Forward<PieSttChar>();
        }
    }
}
