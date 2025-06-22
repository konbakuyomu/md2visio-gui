using md2visio.struc.graph;
using Microsoft.Office.Interop.Visio;

namespace md2visio.struc.figure
{
    internal interface INode
    {
        string ID { get; set; }
        string Label { get; set; }
        Shape? VisioShape { get; set; }
        Container Container { get; set; }
        List<GEdge> InputEdges { get; }
        List<GEdge> OutputEdges { get; }
    }
}
