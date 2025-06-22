using md2visio.struc.graph;
using md2visio.vsdx.@base;

namespace md2visio.vsdx
{
    internal class VBuilderG(Graph figure) : VFigureBuilder<Graph>(figure)
    {
        override protected void ExecuteBuild() 
        {
            new VDrawerG(figure, VisioApp).Draw();
        }        
    }
}
