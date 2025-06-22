using md2visio.mermaid.cmn;
using System.Text;

namespace md2visio.mermaid.pie
{
    internal class PieSttTuple : SynState
    {
        public override SynState NextState()
        {
            if(Buffer.Length > 0) Restore(Buffer.Length).ClearBuffer();
            if (!IsTupleLine(Ctx)) throw new SynException("expected a tuple");

            string tuple = ExpectedGroups[0].Value.Trim();
            int index = tuple.LastIndexOf(':');
            AddToCompo(tuple.Substring(0, index));
            AddToCompo(tuple.Substring(index + 1, tuple.Length - index - 1));

            return ClearBuffer().Save(tuple).Forward<PieSttChar>();
        }

        void AddToCompo(string compo)
        {
            string trimed = TrimItem(compo);
            if (string.IsNullOrEmpty(trimed)) throw new SynException("data item can't be empty", Ctx);
            AddCompo(trimed);
        }

        string TrimItem(string item)
        {
            StringBuilder sb = new StringBuilder(item.Trim());
            if (sb[0] ==  '"') sb.Remove(0, 1);
            if (sb.Length > 0 && sb[sb.Length - 1] == '"') sb.Remove(sb.Length-1, 1);

            return sb.ToString();
        }

        public static bool IsTupleLine(SynContext ctx)
        {
            return ctx.Expect(@"[^\n]+:[^:\n""]+(?=\n)");
        }
    }
}
