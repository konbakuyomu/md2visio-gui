using md2visio.mermaid.cmn;
using md2visio.struc.figure;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph
{
    internal class GSttLinkLabel : SynState
    {
        public override SynState NextState()
        {
            // [<ox]--LINK_LABEL or [<ox]==LINK_LABEL
            if (!IsLinkText(Ctx)) throw new SynException("expected link text", Ctx);

            return Create(Ctx);
        }

        public static SynState Create(SynContext ctx)
        {
            if (!IsLinkText(ctx)) return EmptyState.Instance;

            Group textGroup = ctx.TestGroups["text"];
            string text = textGroup.Value.Trim();

            SynState state = new GSttLinkLabel();
            state.Ctx = ctx;
            state.Save(text).Slide(textGroup.Index + textGroup.Length);

            return state.Forward<GSttLinkEnd>();
        }

        public static bool IsLinkText(SynContext ctx)
        {
            if (ctx.LastState() is not GSttLinkStart) return false;

            string tag = ctx.LastState().Fragment.Substring(1, 1);
            if (!ctx.Test(@$"(?s)^(?<text>.+?)(?=[<xo]?{tag}{{2,}}[{tag}ox>])"))
                return false;

            Group textGroup = ctx.TestGroups["text"];
            string text = textGroup.Value.Trim();
            if (text.Length == 0) return false;

            return true;
        }
    }
}
