using md2visio.struc.journey;
using md2visio.vsdx.@base;

namespace md2visio.vsdx
{
    internal class VBuilderJo(Journey figure) : VFigureBuilder<Journey>(figure)
    {
        protected override void ExecuteBuild()
        {
            new VDrawerJo(figure, VisioApp).Draw();
        }
    }
}
