using md2visio.mermaid.graph;
using md2visio.mermaid.graph.@internal;

namespace md2visio.mermaid.cmn
{
    internal class SttWordFlag : SynState
    {
        public override SynState NextState()
        {
            if (!string.IsNullOrWhiteSpace(Buffer)) return Forward<SttFigureType>();
            if(Peek() == "\n") return Forward<SttFinishFlag>();

            return SlideSpaces().Forward<SttIntro>();
        }
    }
}
