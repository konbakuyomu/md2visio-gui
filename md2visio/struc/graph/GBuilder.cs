using md2visio.mermaid.cmn;
using md2visio.mermaid.graph;
using md2visio.struc.figure;

namespace md2visio.struc.graph
{
    internal class GBuilder : FigureBuilder
    {
        static List<GNode> EmptyList = new List<GNode>();

        Graph graph = new Graph();
        Stack<Graph> stack = new Stack<Graph>();
        List<GNode> fromNodes = EmptyList, toNodes = EmptyList;
        GEdge edge = Empty.Get<GEdge>();

        public GBuilder(SttIterator iter) : base(iter)
        {

        }

        public override void Build(string outputFile)
        {
            while (iter.HasNext())
            {
                SynState cur = iter.Next();
                if (cur is SttMermaidStart) stack.Clear();
                else if (cur is SttMermaidClose) { Output(outputFile); break; }
                else if (cur is GSttKeyword) BuildKeyword();
                else if (cur is GSttText)
                {
                    List<GNode> nodes = GatherNodes();
                    if (fromNodes.Count == 0) fromNodes = nodes;
                    else if (edge.IsEmpty()) fromNodes = nodes;
                    else
                    {
                        toNodes = nodes;
                        ConnectNodes();
                    }
                }
                else if (cur is GSttLinkStart)      BuildEdge();
                else if (cur is GSttNoLabelLink)    BuildEdge();
                else if (cur is SttComment)     { graph.Config.LoadUserDirectiveFromComment(cur.Fragment); }
                else if (cur is SttFrontMatter) { graph.Config.LoadUserFrontMatter(cur.Fragment); }
                else if (cur is SttFinishFlag)  { }
            }
        }

        void ConnectNodes()
        {
            if (edge.IsEmpty()) return;

            foreach (GNode from in fromNodes)
            {
                foreach (GNode to in toNodes)
                {
                    edge.Clone().Connect(from, to);
                }
            }

            edge = Empty.Get<GEdge>();
            fromNodes = toNodes;
            toNodes = EmptyList;
        }

        List<GNode> GatherNodes(List<GNode>? nodes = null)
        {
            SynState cur = iter.Current;
            if (cur is not GSttText) throw new SynException("expected graph node", iter);

            if (nodes == null) nodes = new List<GNode>();

            Graph graph = SuperContainer();
            GNode node = graph.RetrieveNode<GNode>(cur.Fragment);
            nodes.Add(node);
            graph.AddInnerNode(node);

            // shape
            SynState next = iter.PeekNext();
            if (next is GSttExtendShape)
            {
                iter.Next();
                node.NodeShape = GNodeShape.CreateExtend(next.Fragment);
            }
            else if (next is GSttPaired)
            {
                iter.Next();
                string start = next.GetPart("start"),
                    mid = next.GetPart("mid"),
                    close = next.GetPart("close");
                node.NodeShape = GNodeShape.CreatePaired(start, mid, close);
            }
            // &
            else if (next is GSttAmp)
            {
                iter.Next();
                iter.Next();
                GatherNodes(nodes);
            }

            return nodes;
        }

        GEdge BuildEdge()
        {
            SynState state = iter.Current;
            edge = Empty.Get<GEdge>();
            if (state is GSttLinkStart)
            {
                edge = new GEdge();
                edge.StartTag = state.Fragment;

                if (iter.Next() is not GSttLinkLabel) throw new SynException("expected link label", iter);
                edge.Text = iter.Current.Fragment;

                if (iter.Next() is not GSttLinkEnd) throw new SynException("expected link end", iter);
                edge.EndTag = iter.Current.Fragment;
            }
            else if (state is GSttNoLabelLink)
            {
                edge = new GEdge();
                edge.StartTag = state.Fragment;
                edge.EndTag = state.Fragment;

                if (iter.PeekNext() is GSttPipedLinkText)
                {
                    edge.Text = iter.Next().Fragment;
                }
            }

            return edge;
        }

        void BuildKeyword()
        {
            SynState sttNext = iter.PeekNext();
            string frag = iter.Current.Fragment;

            if (frag == "graph" || frag == "flowchart")
            {                
                graph.SetParam(iter.Next().CompoList);
                stack.Push(graph);
            }
            else if (frag == "subgraph")
            {
                Graph graph = SuperContainer();
                GSubgraph subgraph = new GSubgraph(graph);
                subgraph.SetParam(iter.Next().CompoList);

                // add to parent graph
                graph.AddSub(subgraph);
                stack.Push(subgraph);
            }
            else if (frag == "end")
            {
                if (stack.Count == 0) throw new SynException("expected 'graph', 'flowchart' or 'subgraph'", iter);
                stack.Pop();
            }
            else if (frag == "direction")
            {
                if (sttNext is not GSttKeywordParam) throw new SynException("expected keyword param", iter);

                stack.First().Direction = iter.Next().Fragment;
            }
            // TODO
            else if (frag == "click") { }
            else if (frag == "style") { }
            else if (frag == "linkStyle") { }
            else if (frag == "class") { }
            else if (frag == "classDef") { }
        }

        Graph SuperContainer()
        {
            if (stack.Count == 0) throw new SynException("expected a graph or subgraph", iter);

            return stack.First();
        }

        void Output(string outputFile)
        {
            if (stack.Count == 0) return;

            Container nc = stack.First();
            if (nc is Figure) nc.DownCast<Figure>().ToVisio(outputFile);
            stack.Clear();
        }

    }
}
