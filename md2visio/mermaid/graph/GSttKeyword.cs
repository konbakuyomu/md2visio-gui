using md2visio.mermaid.cmn;
using md2visio.mermaid.graph.@internal;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph
{
    internal class GSttKeyword : SynState
    {
        public override SynState NextState()
        {
            if (!IsKeyword(Buffer)) throw new SynException("syntax error", Ctx);
            if (Buffer == "direction")
            {
                if (!GSttKeywordParam.HasParam(Ctx)) return Forward<GSttText>();
            }

            Save(Buffer).ClearBuffer().SlideSpaces();
            if (GSttKeywordParam.HasParam(Ctx)) return Forward<GSttKeywordParam>();

            return Forward<GSttChar>();
        }

        static Regex regKW =
            new("^(graph|flowchart|subgraph|end|style|linkStyle|class|classDef|direction|click)$", RegexOptions.Compiled);
        public static bool IsKeyword(string word)
        {
            return regKW.IsMatch(word);
        }

        static Regex regDir =
            new("^(LR|RL|TD|TB|BT)$", RegexOptions.Compiled);
        public static bool IsDirection(string word)
        {
            return regDir.IsMatch(word);
        }
    }
}
