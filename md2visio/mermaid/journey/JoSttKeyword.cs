using md2visio.mermaid.cmn;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.journey
{
    internal class JoSttKeyword : SynState
    {
        public override SynState NextState()
        {
            Save(Buffer).ClearBuffer();
            if (JoSttKeywordParam.HasParam(Ctx)) return Forward<JoSttKeywordParam>();
            else return Forward<JoSttChar>();
        }

        public static bool IsKeyword(SynContext ctx)
        {
            return Regex.IsMatch(ctx.Cache.ToString(), "^(journey|title|section)$");
        }
    }
}
