namespace md2visio.mermaid.cmn
{
    internal class SttIntro: SynState
    {
        public override SynState NextState()
        {
            string? next = Ctx.Peek();
            if (next == null) return EndOfFile;

            if (next == "-") { return Forward<SttFrontMatter>(); }
            if (next == "%") { return Forward<SttPercent>(); }
            if (next == " ") { return Forward<SttWordFlag>(); }
            if (next == "\t") { return Forward<SttWordFlag>(); }
            if (next == "\n") { return Forward<SttWordFlag>(); }

            if (Unexpected(next)) throw new SynException($"unexpected '{next}'", Ctx);

            return Take().Forward<SttIntro>();
        }

        bool Unexpected(string next)
        {
            switch(next)
            {
                case ";":
                case "@":
                case "~":
                case "{":
                case "[":
                case "(":
                case "<":
                case "`":
                case "\"":
                case "=":
                case "&":
                case "|":
                case ")":
                case "}":
                case "]":
                case ">": return true;
            }
            return false;
        }
    }
}
