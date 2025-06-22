using md2visio.struc.figure;
using md2visio.vsdx;

namespace md2visio.struc.xy
{
    internal class XyChart : Figure
    {
        public XyAxis XAxis { get; set; } = Empty.Get<XyAxis>();
        public XyAxis YAxis { get; set; } = Empty.Get<XyAxis>();
        public MmdJsonArray Bar {  get; set; } = new MmdJsonArray();
        public MmdJsonArray Line { get; set; } = new MmdJsonArray();

        public XyChart() {
            
        }

        public override void ToVisio(string path)
        {
            new VBuilderXy(this).Build(path);
        }
    }
}
