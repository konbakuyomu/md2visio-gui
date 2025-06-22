using md2visio.mermaid.cmn;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.xy
{
    internal class XySttKeyword : SynState
    {
        public override SynState NextState()
        {
            if (!IsKeyword(Ctx)) throw new SynException($"unknown keyword '{Buffer}'", Ctx);

            Save(Buffer).ClearBuffer();
            if (XySttKeywordParam.HasParam(Ctx)) return Forward<XySttKeywordParam>();
            else return Forward<XySttChar>();
        }

        public static bool IsKeyword(SynContext ctx)
        {
            return Regex.IsMatch(ctx.Cache.ToString(), "^(xychart-beta|xychart|title|x-axis|y-axis|line|bar)$");
        }
    }
}
