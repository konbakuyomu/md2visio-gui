using md2visio.mermaid.cmn;
using md2visio.mermaid.graph.@internal;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.graph
{
    internal class GSttText : SynState
    {
        Regex regNodeString = new(@"^[\^=\|><]|~{1,2}$", RegexOptions.Compiled);

        public override SynState NextState()
        {
            if (!Ctx.FindContainerFrag("graph|flowchart|subgraph").Success)
                throw new SynException("syntax error", Ctx);

            string text = Buffer;
            ClearBuffer().SlideSpaces();
            if (string.IsNullOrWhiteSpace(text)) return Forward<GSttChar>();

            if (regNodeString.IsMatch(text))
                throw new SynException("unexpected node string", Ctx);

            return Save(text).Forward<GSttChar>();
        }
    }
}
