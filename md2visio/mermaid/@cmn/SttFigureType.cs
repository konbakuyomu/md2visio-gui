using md2visio.mermaid.graph;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.cmn
{
    internal class SttFigureType : SynState
    {
        public static readonly string Supported = 
            "graph|flowchart|sequenceDiagram|classDiagram|stateDiagram|stateDiagram-v2|" +
            "erDiagram|journey|gantt|pie|quadrantChart|requirementDiagram|gitGraph|C4Context|mindmap|" +
            "timeline|zenuml|sankey|sankey-beta|xychart|xychart-beta|block|block-beta|packet|packet-beta|" +
            "kanban|architecture|architecture-beta";

        Dictionary<string, Type> typeMap = TypeMap.KeywordMap;

        public override SynState NextState()
        {
            string kw = Buffer.Trim();
            if (!IsFigure(kw)) throw new SynException($"unknown graph type {kw}", Ctx);
            if (!typeMap.ContainsKey(kw)) throw new NotImplementedException($"keyword '{kw}' not implemented");

            return Forward(typeMap[kw]);
        }

        public static bool IsFigure(string word)
        {
            return Regex.IsMatch(word, $"^({Supported})$");
        }
    }
}
