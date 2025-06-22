using md2visio.mermaid.cmn;
using md2visio.struc.figure;
using Microsoft.Office.Interop.Visio;
using System.Text;

namespace md2visio.struc.graph
{
    internal class GSubgraph : Graph, INode
    {
        GLabel label = new GLabel(string.Empty);
        List<GNode>? allGroupedNodes = null;
        GNode borderNode = Empty.Get<GBorderNode>();

        public GSubgraph() { }

        public GNode BorderNode { get => borderNode; }
        public List<GNode> AllGroupedNodes { get => GetAllGroupedNodes(); }

        public string ID { get => $"_{borderNode.ID}_"; set => SetID(value); }
        public string Label
        {
            get { return label.ToString(); }
            set { label.Content = value; }
        }
        public Container Container { 
            get => Parent ?? Empty.Get<Container>(); 
            set { SetContainer(value); } 
        }

        public GSubgraph(Graph parent) 
        {
            Parent = parent;
        }
        public Shape? VisioShape { 
            get => borderNode.VisioShape; 
            set => borderNode.VisioShape = value; 
        }

        public List<GEdge> InputEdges { get; set; } = new List<GEdge>();

        public List<GEdge> OutputEdges { get; set; } = new List<GEdge>();

        public override void SetParam(CompoDict valueList)
        {
            StringBuilder t = new StringBuilder();
            foreach (var item in valueList.Values())
            {
                if (item.IsPaired()) Label = item.Value;
                else if (item.IsText()) t.Append(item.Value);
                else if (item.IsQuoted()) t.Append(item.Value);
            }
            if (t.Length == 0) throw new SynException("expected subgraph ID");
            ID = t.ToString(); 
        }
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            List<INode> nodes = groupedNodes.Values.ToList();
            foreach(INode node in nodes)
            {
                if (node == nodes.First())  sb.Append(": ");
                if (node.Container == this) sb.Append(node.ID);
                if (node != nodes.Last())   sb.Append(", ");
            }
            sb.Insert(0, ID);
            return sb.ToString();
        }

        void SetContainer(Container value)
        {
            if (Parent != null)
            {
                Parent.GroupedNodes.Remove(ID);
            }
            Parent = value;
            Parent.AddInnerNode(this);
        }

        void SetID(string value)
        {
            if (nodeDict.ContainsKey(value))
            {
                nodeDict[value].Container.GroupedNodes.Remove(value);
                nodeDict.Remove(value);
            }

            borderNode = RetrieveNode<GBorderNode>(value);
            borderNode.Container = this;

            if (label.Content.Length == 0) Label = value;
        }

        List<GNode> GetAllGroupedNodes()
        {
            if(allGroupedNodes != null) return allGroupedNodes;

            allGroupedNodes = new(groupedNodes.Values.OfType<GNode>())
            {
                borderNode
            };
            foreach(GSubgraph sub in Subgraphs)
            {
                allGroupedNodes.AddRange(sub.GetAllGroupedNodes());
            }

            return allGroupedNodes;
        }


    }

}
