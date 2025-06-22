using md2visio.mermaid.cmn;

namespace md2visio.mermaid.journey
{
    internal class JoSttWord : SynState
    {
        public override SynState NextState()
        {
            if (string.IsNullOrEmpty(Buffer)) return SlideSpaces().Forward<JoSttChar>();

            if(JoSttKeyword.IsKeyword(Ctx)) return Forward<JoSttKeyword>();
            return Take().Forward<JoSttChar>();
        }
    }
}
