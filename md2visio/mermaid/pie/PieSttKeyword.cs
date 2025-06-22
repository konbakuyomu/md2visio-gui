using md2visio.mermaid.cmn;
using md2visio.mermaid.journey;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.pie
{
    internal class PieSttKeyword : SynState
    {
        public override SynState NextState()
        {
            if (!IsKeyword(Ctx)) throw new SynException($"unknown keyword '{Buffer}'", Ctx);

            Save(Buffer).ClearBuffer();
            if (PieSttKeywordParam.HasParam(Ctx)) return Forward<PieSttKeywordParam>();
            else return Forward<PieSttChar>();
        }

        public static bool IsKeyword(SynContext ctx)
        {
            return Regex.IsMatch(ctx.Cache.ToString(), "^(pie|title)$");
        }
    }
}
