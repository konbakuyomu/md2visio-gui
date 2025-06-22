using md2visio.struc.figure;
using md2visio.struc.graph;
using Microsoft.Office.Interop.Visio;

namespace md2visio.struc.pie
{
    internal class PieDataItem : INode
    {
        public string ID { get; set; } = string.Empty;
        public string Label { get => ID; set => ID = value; }
        public Shape? VisioShape { get; set; }
        public Container Container { get; set; } = Empty.Get<Container>();
        public double Data {  get; set; }

        public List<GEdge> InputEdges => throw new NotImplementedException();

        public List<GEdge> OutputEdges => throw new NotImplementedException();
    }
}
