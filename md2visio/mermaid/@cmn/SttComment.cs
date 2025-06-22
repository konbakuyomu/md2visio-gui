using md2visio.mermaid.graph;

namespace md2visio.mermaid.cmn
{
    internal class SttComment : SynState
    {
        public override SynState NextState()
        {
            if (Expect(@"(?s)(?<cmnt>%%[^\n]+?(%%)?(?=\n))")) 
                return Save(ExpectedGroup("cmnt")).Forward<SttCtxChar>();

            return Take(2).Forward<SttCtxChar>();
        }
    }
}
