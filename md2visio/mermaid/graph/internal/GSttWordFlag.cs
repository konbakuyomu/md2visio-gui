using md2visio.mermaid.cmn;

namespace md2visio.mermaid.graph.@internal
{
    internal class GSttWordFlag : SynState
    {
        public override SynState NextState()
        {
            // a word flag may make a node string
            // left
            if (!string.IsNullOrWhiteSpace(Buffer)) { return Forward<GSttWord>(); }

            // right            
            string? next = Peek();
            if (next == "\n") { return Forward<SttFinishFlag>(); }

            next = SlideSpaces().Peek();
            if (next == ";") { return Forward<SttFinishFlag>(); }
            if (next == "`") { return Forward<GSttBackQuote>(); }
            if (next == "=") { return Forward<GSttNoLabelLink>(); }
            if (next == ">") { return Forward<GSttNoLabelLink>(); }
            if (next == "<") { return Forward<GSttLinkStart>(); }
            if (next == "{") { return Forward<GSttPaired>(); }
            if (next == "[") { return Forward<GSttPaired>(); }
            if (next == "(") { return Forward<GSttPaired>(); }
            if (next == "@") { return Forward<GSttExtendShape>(); }
            if (next == "%") { return Forward<SttPercent>(); }
            if (next == "~") { return Forward<GSttTilde>(); }

            return Forward<GSttChar>();
        }
    }
}
