using md2visio.struc.figure;
using md2visio.vsdx;

namespace md2visio.struc.pie
{
    internal class Pie : Figure
    {
        public bool ShowData { get; set; } = false;

        public Pie() { 
        }

        public double TotalNum()
        {
            double total = 0;
            foreach (INode item in InnerNodes.Values)
            {
                total += ((PieDataItem)item).Data;
            }
            return total;
        }

        public override void ToVisio(string path)
        {
            new VBuilderPie(this).Build(path);
        }
    }
}
