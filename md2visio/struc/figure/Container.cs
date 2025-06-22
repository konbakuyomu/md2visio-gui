using md2visio.struc.graph;
using System.Text.RegularExpressions;

namespace md2visio.struc.figure
{
    internal class Container
    {
        string direction = "LR";
        GrowthDirection shift;
        Container? parent = null;
        List<Container>? ancestors = null;
        
        protected Dictionary<string, INode> nodeDict = new Dictionary<string, INode>();
        protected Dictionary<string, INode> innerNodes = new Dictionary<string, INode>();
        protected Dictionary<string, INode> groupedNodes = new Dictionary<string, INode>(); // nodes grouped by a subgraph

        public string Direction
        {
            get { return direction; }
            set { if (IsDirectionFragment(value)) { direction = value.Trim(); shift.Decide(this); } }
        }
        public Dictionary<string, INode> InnerNodes { get { return innerNodes; } }
        public Dictionary<string, INode> NodeDict { get { return nodeDict; } }
        public GrowthDirection GrowthDirect { get { return shift; } }
        public Container? Parent
        {
            get => parent;
            set { if (value != null) nodeDict = value.nodeDict; parent = value; }
        }
        public List<Container> Ancestors { get => GetAncestorList(); }
        public Dictionary<string, INode> GroupedNodes { get => groupedNodes; }
        public Container()
        {
            shift = new GrowthDirection(this);
        }       

        public virtual T RetrieveNode<T>(string id) where T : INode, new()
        {
            if (nodeDict.ContainsKey(id)) return (T)nodeDict[id];

            T node = new T();
            node.Container = this;
            node.ID = id;

            AddNodeToDict(node);
            return node;
        }

        public void AddInnerNode(INode node)
        {
            if (innerNodes.ContainsKey(node.ID)) return;            

            innerNodes.Add(node.ID, node);
            AddNodeToDict(node);

            // judge container
            if (node.Container.IsEmpty()) node.Container = this;
            else if (node.Container.GetType() == typeof(Graph))
            {
                if (GetType() == typeof(GSubgraph))
                {
                    node.Container.groupedNodes.Remove(node.ID);
                    node.Container = this;
                }
            }
            else if (node.Container.GetType() == typeof(GSubgraph))
            {
                if (Ancestors.Contains(node.Container))
                {
                    node.Container.groupedNodes.Remove(node.ID);
                    node.Container = this;
                }
            }             

            if(node.Container == this && !groupedNodes.ContainsKey(node.ID)) 
                groupedNodes.Add(node.ID, node);
        }

        public T DownCast<T>() where T : Container
        {
            if (this is T) return (T)this;

            throw new InvalidCastException();
        }

        protected void AddNodeToDict(INode node)
        {
            if (!nodeDict.ContainsKey(node.ID))
                nodeDict.Add(node.ID, node);
        }

        static public bool IsDirectionFragment(string fragment)
        {
            return Regex.IsMatch(fragment.Trim(), "^(LR|RL|TB|TD|BT)$");
        }

        List<Container> GetAncestorList()
        {
            if (ancestors != null) return ancestors;

            ancestors = new List<Container>();
            Container? parent = Parent;
            while (parent != null)
            {
                ancestors.Add(parent);
                parent = parent.Parent;
            }

            return ancestors;
        }
    }
}
