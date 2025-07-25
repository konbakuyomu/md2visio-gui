using md2visio.mermaid.cmn;
using md2visio.mermaid.graph.@internal;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph
{
    internal class GSttPaired : SynState
    {
        public override SynState NextState()
        {
            if (!IsPaired(Ctx)) throw new SynException("expected pair start", Ctx);

            string pairStart = Ctx.TestGroups["start"].Value;
            if (!IsShapeStart(pairStart)) throw new SynException("expected shape start", Ctx);

            if (IsPaired(pairStart, Ctx))
            {
                string mid = Ctx.TestGroups["mid"].Value;
                string close = Ctx.TestGroups["close"].Value;
                if (mid.Length == 0) throw new SynException("expected pair label", Ctx);

                string pair = Ctx.TestGroups[0].Value;
                AddCompo("start", pairStart);
                AddCompo("mid", mid);
                AddCompo("close", close);
                return Save(pair).Slide(Ctx.TestGroups[0].Length).Forward<GSttChar>();
            }

            throw new SynException("expected pair close", Ctx);
        }

        public static bool IsPaired(SynContext ctx)
        {
            if (!ctx.Test(@"(?<start>[\[\{\(>]+)")) return false;
            return IsPaired(ctx.TestGroups["start"].Value, ctx);
        }

        public static bool IsPaired(string pairStart, SynContext ctx)
        {
            if (MmdPaired.IsPaired(pairStart, ctx.Incoming))
            {
                ctx.TestGroups = MmdPaired.Groups;
                return true;
            }
            return false;
        }

        static Regex regShapeStart =
            new(@"^(>|\[{1,2}|\{{1,2}|\({1,3}|\[\(|\[\\|\[/|\(\[)$", RegexOptions.Compiled);
        bool IsShapeStart(string fragment)
        {
            return regShapeStart.IsMatch(fragment);
        }

        static Regex regShapeClose =
            new(@"^(\]{1,2}|\}{1,2}|\){1,3}|\)\]|\\\]|/\]|\]\)|\]\))$", RegexOptions.Compiled);
        bool IsShapeClose(string fragment)
        {
            return regShapeClose.IsMatch(fragment);
        }
    }
}
