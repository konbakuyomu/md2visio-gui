using md2visio.mermaid.cmn;

namespace md2visio.mermaid.pie
{
    internal class PieSttWord : SynState
    {
        public override SynState NextState()
        {
            if (string.IsNullOrEmpty(Buffer)) return SlideSpaces().Forward<PieSttChar>();

            if (PieSttKeyword.IsKeyword(Ctx)) return Forward<PieSttKeyword>();
            return Take().Forward <PieSttChar>();
        }
    }
}
