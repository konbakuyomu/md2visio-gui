using md2visio.mermaid.cmn;
using md2visio.struc.figure;
using System.Drawing;

namespace md2visio.struc.xy
{
    internal class XyAxis
    {
        public string Title { get; set; } = string.Empty;
        public SizeF Range { get; set; } = SizeF.Empty;
        public MmdJsonArray Values { get; set;} = new MmdJsonArray();

        public XyAxis() { }

        public XyAxis(CompoDict dict)
        {
            Title = dict["title"] ?? string.Empty;
            if (dict.ContainsKey("s") && dict.ContainsKey("e"))
                InitRange(dict.Get("s"), dict.Get("e"));
            if (dict.ContainsKey("vals"))
                Values.Load(dict.Get("vals"));
        }

        void InitRange(string rangeStart, string rangeEnd)
        {
            float start, end;
            if(float.TryParse(rangeStart, out start) &&
                float.TryParse(rangeEnd, out end))
                Range = new SizeF(start, end);
        }
    }
}
