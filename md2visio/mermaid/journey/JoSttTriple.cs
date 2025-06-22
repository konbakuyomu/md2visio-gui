using md2visio.mermaid.cmn;

namespace md2visio.mermaid.journey
{
    internal class JoSttTriple : SynState
    {
        public override SynState NextState()
        {
            if(!IsTripleLine(Ctx)) throw new SynException("expected a triple", Ctx);

            string triple = ExpectedGroups[0].Value.Trim();
            foreach (string s in triple.Split(":"))
            {
                string trimed = s.Trim();
                if (string.IsNullOrEmpty(trimed)) throw new SynException("task item can't be empty", Ctx);
                AddCompo(trimed);
            }

            return Save(triple).Forward<JoSttChar>();
        }

        public static bool IsTripleLine(SynContext ctx)
        {
            return ctx.Expect(@"([^:\n]+:){2}[^:\n]+(?=\n)");
        }
    }
}
