namespace md2visio.mermaid.cmn
{
    /**
     * ---
     * title: This is title
     * config:
     *     ...
     * ---
     */
    internal class SttFrontMatter : SynState
    {
        public override SynState NextState()
        {
            if (!IsConfig(Ctx)) return Take().Forward<SttCtxChar>();
            if (!Until(@"(?s)---\s*(?<cfg>.*?)\n\s*---\s*(?=\n)")) throw new SynException("expected '---'", Ctx);

            Save(ExpectedGroup("cfg"));
            return ClearBuffer().Forward<SttIntro>();
        }

        public static bool IsConfig(SynContext ctx)
        {
            return ctx.Test(@"^(?<cstart>---\s*\n)");
        }
    }
}
