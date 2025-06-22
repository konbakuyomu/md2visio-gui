using md2visio.mermaid.cmn;

namespace md2visio.mermaid.graph.@internal
{
    internal class GSttChar : SynState
    {
        public override SynState NextState()
        {
            string? next = Ctx.Peek();
            if (next == null) return EndOfFile;

            if (next == ";") { return Forward<GSttWordFlag>(); }
            if (next == "\n") { return Forward<GSttWordFlag>(); }
            if (next == "\t") { return Forward<GSttWordFlag>(); }
            if (next == " ") { return Forward<GSttWordFlag>(); }
            if (next == "@") { return Forward<GSttWordFlag>(); }
            if (next == "~") { return Forward<GSttWordFlag>(); }
            if (next == "{") { return Forward<GSttWordFlag>(); }
            if (next == "[") { return Forward<GSttWordFlag>(); }
            if (next == "(") { return Forward<GSttWordFlag>(); }
            if (next == "<") { return Forward<GSttWordFlag>(); }
            if (next == "`") { return Forward<GSttBackQuote>(); }
            if (next == "\"") { return Forward<GSttQuoted>(); }
            if (next == "-") { return Forward<GSttMinus>(); }
            if (next == "=") { return Forward<GSttEqual>(); }
            if (next == "&") { return Forward<GSttAmp>(); }
            if (next == "%") { return Forward<SttPercent>(); }
            if (next == "|") { return Forward<GSttPipedLinkText>(); }
            if (next == ")") { throw new SynException("unexpected ')'", Ctx); }
            if (next == "}") { throw new SynException("unexpected '}'", Ctx); }
            if (next == "]") { throw new SynException("unexpected ']'", Ctx); }
            if (next == ">") { throw new SynException("unexpected '>'", Ctx); }

            return Take().Forward<GSttChar>();
        }
    }
}
