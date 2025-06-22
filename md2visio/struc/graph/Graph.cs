using md2visio.mermaid.cmn;
using md2visio.struc.figure;
using md2visio.vsdx;

namespace md2visio.struc.graph
{
    internal class Graph : Figure
    {
        Dictionary<string, GSubgraph> subgraphDict = new Dictionary<string, GSubgraph>();
        protected LinkedList<GNode> alignedInnerNodes = new LinkedList<GNode>();
        protected LinkedList<GNode> alignedGroupedNodes = new LinkedList<GNode>();

        public List<GSubgraph> Subgraphs { get { return subgraphDict.Values.ToList(); } }

        public Graph()
        {
        }

        public LinkedList<GNode> AlignInnerNodes()
        {
            if (innerNodes.Count == 0) return alignedInnerNodes;
            if (alignedInnerNodes.Count > 0) return alignedInnerNodes;

            foreach (GNode node in innerNodes.Values)
            {
                if (alignedInnerNodes.Count == 0) { alignedInnerNodes.AddLast(node); continue; }

                bool sorted = false;
                for (LinkedListNode<GNode>? sortedNode = alignedInnerNodes.Last; sortedNode != null; sortedNode = sortedNode.Previous)
                {
                    GNode? connectionStart = NodeConnectionStart(sortedNode.Value, node);
                    if (connectionStart == sortedNode.Value)
                    {
                        alignedInnerNodes.AddAfter(sortedNode, node);
                        sorted = true;
                        break;
                    }
                }

                if (!sorted) alignedInnerNodes.AddLast(node);
            }
            return alignedInnerNodes;
        }

        public LinkedList<GNode> AlignGroupedNodes()
        {
            if (groupedNodes.Count == 0) return alignedGroupedNodes;
            if (alignedGroupedNodes.Count > 0) return alignedGroupedNodes;

            foreach(GNode node in AlignInnerNodes())
            {
                if(groupedNodes.ContainsKey(node.ID)) 
                    alignedGroupedNodes.AddLast(node);
            }
            return alignedGroupedNodes;
        }

        public static GNode? NodeConnectionStart(LinkedList<GNode> list, GNode node)
        {
            foreach (GNode gNode in list)
            {
                GNode? startNode = NodeConnectionStart(gNode, node);
                if (startNode != null) return startNode;
            }

            return null;
        }

        public static GNode? NodeConnectionStart(GNode nodeA, GNode nodeB)
        {
            if (nodeA.OutputNodes().Contains(nodeB)) return nodeA;
            if (nodeB.OutputNodes().Contains(nodeA)) return nodeB;

            return null;
        }

        public void AddSub(GSubgraph subgraph)
        {
            if (subgraphDict.ContainsKey(subgraph.ID))
            {
                subgraphDict[subgraph.ID] = subgraph;
            }
            else
            {
                subgraphDict.Add(subgraph.ID, subgraph);
                nodeDict.Add(subgraph.ID, subgraph);
            }
        }

        public (GNode? linkedNode, RelativePos nodePos) LinkedNode(LinkedList<GNode> nodes2test)
        {
            foreach (GNode innerNode in innerNodes.Values)
            {
                foreach (GNode testNode in nodes2test)
                {
                    GNode? start = NodeConnectionStart(innerNode, testNode);
                    if (start == null) continue;
                    else if (start == innerNode) return (testNode, RelativePos.TAIL);
                    else return (testNode, RelativePos.FRONT);
                }
            }

            return (null, RelativePos.TAIL);
        }

        (GNode? linkedNode, RelativePos nodePos) LinkedNode(LinkedList<GNode> drawList, Graph graph, LinkedList<GNode> drawnNodes)
        {
            foreach (GNode node2draw in drawList)
            {
                foreach (GSubgraph subGraph in graph.Subgraphs)
                {
                    LinkedList<GNode> subNodeList = subGraph.AlignInnerNodes();
                    (GNode? linkedNode, RelativePos nodePos) = NodeLinkedToSubgraph(node2draw, subNodeList, drawnNodes);
                    if (linkedNode == null) continue;

                    return (node2draw, nodePos);
                }
            }

            foreach (GSubgraph subGraph in graph.Subgraphs)
            {
                (GNode? linkedNode, RelativePos nodePos) = LinkedNode(drawList, subGraph, drawnNodes);
                if (linkedNode == null) continue;

                return (linkedNode, nodePos);
            }


            foreach (GNode node2draw in drawList)
            {
                (GNode? linkedNode, RelativePos nodePos) = NodeLinkedToSubgraph(node2draw, drawList, drawnNodes);
                if (linkedNode == null) continue;

                return (node2draw, nodePos);
            }

            return (null, RelativePos.TAIL);
        }

        (GNode? linkedNode, RelativePos nodePos) NodeLinkedToSubgraph(GNode node2draw, LinkedList<GNode> compareList, LinkedList<GNode> excludeNodes)
        {
            foreach (GNode cmpNode in compareList)
            {
                GNode? start = NodeConnectionStart(node2draw, cmpNode);
                if (start == null) continue;
                if (excludeNodes.Contains(node2draw)) continue;

                excludeNodes.AddLast(node2draw);
                if (start == node2draw) return (node2draw, RelativePos.FRONT);
                else return (node2draw, RelativePos.TAIL);
            }
            return (null, RelativePos.TAIL);
        }

        public (GNode? linkedNode, RelativePos nodePos) NodesHavingInput(LinkedList<GNode> drawnNodes)
        {
            LinkedList<GNode> nodes2draw = AlignGroupedNodes(); 
            return LinkedNode(nodes2draw, this, drawnNodes);
        }

        public virtual void SetParam(CompoDict list)
        {
            string frag = list.Entire;
            if (!string.IsNullOrEmpty(frag)) Direction = frag;
        }

        public override void ToVisio(string path)
        {
            new VBuilderG(this).Build(path);
        }

    }
}
