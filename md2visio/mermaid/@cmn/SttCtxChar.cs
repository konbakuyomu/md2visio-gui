using md2visio.mermaid.graph.@internal;
using md2visio.mermaid.pie;

namespace md2visio.mermaid.cmn
{
    internal class SttCtxChar : SynState
    {
        Dictionary<string, Type> typeMap = TypeMap.CharMap;

        public override SynState NextState()
        {
            (bool success, string graph) = Ctx.FindContainerFrag(SttFigureType.Supported);
            if (success) return Forward(typeMap[graph]);
            else return Forward<SttIntro>();
        }
    }
}
