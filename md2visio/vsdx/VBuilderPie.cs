using md2visio.struc.pie;
using md2visio.vsdx.@base;

namespace md2visio.vsdx
{
    internal class VBuilderPie(Pie figure) : VFigureBuilder<Pie>(figure)
    {
        protected override void ExecuteBuild()
        {
            new VDrawerPie(figure, VisioApp).Draw();
        }

    }
}
