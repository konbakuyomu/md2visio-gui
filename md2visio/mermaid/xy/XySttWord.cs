using md2visio.mermaid.cmn;

namespace md2visio.mermaid.xy
{
    internal class XySttWord : SynState
    {
        public override SynState NextState()
        {
            if (string.IsNullOrEmpty(Buffer)) return SlideSpaces().Forward<XySttChar>();

            if (XySttKeyword.IsKeyword(Ctx)) return Forward<XySttKeyword>();
            return Take().Forward<XySttChar>();
        }
    }
}
