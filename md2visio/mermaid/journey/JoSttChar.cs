using md2visio.mermaid.cmn;

namespace md2visio.mermaid.journey
{
    internal class JoSttChar : SynState
    {
        public override SynState NextState()
        {
            string? next = Ctx.Peek();
            if (next == null) return EndOfFile;

            if (next == "%")  return Forward<SttPercent>(); 
            if (next == ":")  return Restore(Buffer.Length).Forward<JoSttTriple>();
            if (next == " ")  return Forward<JoSttWord>(); 
            if (next == "\t") return Forward<JoSttWord>(); 
            if (next == "\n") return Forward<SttFinishFlag>(); 
            if (next == "`")  return Forward<SttMermaidClose>();

            return Take().Forward<JoSttChar>();
        }


    }
}
