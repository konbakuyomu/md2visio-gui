using md2visio.struc.figure;
using md2visio.struc.packet;
using md2visio.vsdx.@base;

namespace md2visio.vsdx
{
    internal class VBuilderPac(Packet figure) : VFigureBuilder<Packet>(figure)
    {
        protected override void ExecuteBuild()
        {
            VDrawerPac drawer = new VDrawerPac(figure, VisioApp);
            drawer.SortedNodes = OrderInnerNodes();
            drawer.Draw();
        }

        List<INode> OrderInnerNodes()
        {
            List<INode> nodes = figure.InnerNodes.Values.ToList<INode>();
            IComparer<INode> comparer = new PacBitsComparer();
            nodes.Sort(comparer);

            return nodes;
        }
    }

    class PacBitsComparer : IComparer<INode>
    {
        public int Compare(INode? x, INode? y)
        {
            if(x == null || y == null) return 0;
            PacBlock bitsx = (PacBlock) x, 
                bitsy = (PacBlock) y;

            return bitsx.BitStart - bitsy.BitStart;
        }
    }
}
