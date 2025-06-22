using md2visio.mermaid.cmn;
using System.Text;

namespace md2visio.mermaid.packet
{
    internal class PacSttTuple : SynState
    {
        public override SynState NextState()
        {
            if (Buffer.Length > 0) Restore(Buffer.Length).ClearBuffer();
            if (!IsTupleLine(Ctx)) throw new SynException("expected a tuple");

            AddCompo("bits", ExpectedGroups["bits"].Value);
            AddCompo("name", ExpectedGroups["name"].Value);

            return ClearBuffer().Save(ExpectedGroups[0].Value.Trim()).Forward<PaSttChar>();
        }

        public static bool IsTupleLine(SynContext ctx)
        {
            return ctx.Expect(@"^\s*(?<bits>\d+(\s*-\s*\d+)?)\s*:\s*""(?<name>.+)""\s*(?=\n)");
        }
    }
}
