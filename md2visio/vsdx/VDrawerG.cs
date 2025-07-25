using md2visio.mermaid.cmn;
using md2visio.struc.figure;
using md2visio.struc.graph;
using md2visio.vsdx.@base;
using Microsoft.Office.Interop.Visio;

namespace md2visio.vsdx
{
    internal enum RelativePos
    {
        FRONT, TAIL
    }

    internal class VDrawerG : VFigureDrawer<Graph>        
    {
        LinkedList<GNode> drawnList = new LinkedList<GNode>();
        public VDrawerG(Graph figure, Application visioApp): base(figure, visioApp)
        {
        }

        public override void Draw()
        {
            EnsureVisible(); // 确保Visio可见
            PauseForViewing(300); // 给用户时间看到初始状态
            
            DrawNodes(figure);
            PauseForViewing(500); // 节点绘制完成后暂停
            
            DrawEdges(figure);
            PauseForViewing(300); // 边绘制完成后暂停
        }        

        void DrawNodes(Graph graph)
        {
            foreach (GSubgraph subGraph in graph.Subgraphs)
            {
                DrawNodes(subGraph);
            }

            // align grouped nodes
            LinkedList<GNode> nodes2Draw = new LinkedList<GNode>();
            (GNode? linkedNode, RelativePos rpos) = graph.NodesHavingInput(nodes2Draw);
            while (linkedNode != null)
            {
                if (linkedNode.VisioShape == null) break;
                (linkedNode, rpos) = graph.NodesHavingInput(nodes2Draw);
            }

            foreach (GNode node in graph.AlignGroupedNodes()) 
            {
                if(!nodes2Draw.Contains(node)) nodes2Draw.AddLast(node);
            }

            // draw nodes
            if (nodes2Draw.Count == 0) return;

            DrawLinkedNodes(nodes2Draw, graph.GrowthDirect);
            Relocate(nodes2Draw.ToList(), graph.GrowthDirect);

            // border
            if (nodes2Draw.First().Container is GSubgraph)
            {
                GSubgraph subgraph = nodes2Draw.First().Container.DownCast<GSubgraph>();
                DrawSubgraphBorder(subgraph);
                Relocate(subgraph);
            }        
        }

        void DrawLinkedNodes(LinkedList<GNode> nodes, GrowthDirection direct)
        {
            if (nodes.Count == 0) return;

            VBoundary alignBound = Empty.Get<VBoundary>();
            foreach (GNode node in nodes)
            {
                if (node is GBorderNode) continue;

                Shape shape = CreateShape(node);
                PauseForViewing(150); // 每个节点创建后暂停
                
                SetFillForegnd(shape, "config.themeVariables.primaryColor");
                SetLineColor(shape, "config.themeVariables.primaryBorderColor");
                SetTextColor(shape, "config.themeVariables.primaryTextColor");
                shape.CellsU["LineWeight"].FormulaU = "0.75 pt";
                PauseForViewing(100); // 样式设置后暂停
                
                if (alignBound.IsEmpty()) alignBound = new VShapeBoundary(shape);
                else
                {
                    alignBound.Grow(shape, direct, GNode.SPACE);
                    alignBound = new VShapeBoundary(shape);
                }
                drawnList.AddLast(node);
            }
        }

        void DrawEdges(Graph graph)
        {
            List<GEdge> drawnEdges = new List<GEdge>();
            foreach (INode node in graph.NodeDict.Values)
            {
                if (node.VisioShape == null) continue;
                foreach (GEdge edge in node.OutputEdges)
                {
                    if (drawnEdges.Contains(edge) || edge.To.VisioShape == null) continue;

                    Shape shape = CreateEdge(edge);
                    node.VisioShape.AutoConnect(edge.To.VisioShape, VisAutoConnectDir.visAutoConnectDirNone, shape);
                    shape.Delete();
                    drawnEdges.Add(edge);
                    PauseForViewing(100); // 每条边绘制后暂停
                }
            }
        }

        GNode DrawSubgraphBorder(GSubgraph subGraph)
        {
            GNode borderNode = DropSubgraphBorder(subGraph);
            drawnList.AddLast(borderNode);

            return borderNode;
        }

        VBoundary Relocate(GSubgraph subgraph)
        {
            return Relocate(subgraph.AllGroupedNodes, subgraph.Container.GrowthDirect);
        }

        VBoundary Relocate(List<GNode> nodes, GrowthDirection direct)
        {
            VBoundary nodesBound = NodesBoundary(nodes);
            VBoundary relativeBound = NodesBoundary(drawnList.Except(nodes).ToList());
            if (nodes.Count == 0 || relativeBound.Height == 0) return nodesBound;

            VBoundary newBound = new(true);
            bool drawAtTail = IsDrawAtTail(nodes, direct);
            if(direct.H != 0)
            {
                double moveH = 0, moveV = relativeBound.PinY-nodesBound.PinY;
                if (drawAtTail) moveH = relativeBound.Right + GNode.SPACE - nodesBound.Left;
                else moveH = relativeBound.Left - GNode.SPACE - nodesBound.Right;

                foreach (GNode node in nodes) 
                {
                    if (node.VisioShape == null) continue;
                    MoveTo(node.VisioShape, PinX(node.VisioShape)+moveH, PinY(node.VisioShape)+moveV);  
                    newBound.Expand(node.VisioShape);
                }                
            }
            if(direct.V != 0)
            {
                double moveV = 0, moveH = relativeBound.PinX - nodesBound.PinX;
                if (drawAtTail) moveV = relativeBound.Top + GNode.SPACE - nodesBound.Bottom;
                else moveV = relativeBound.Bottom - GNode.SPACE - nodesBound.Top;

                foreach (GNode node in nodes)
                {
                    if (node.VisioShape == null) continue;
                    MoveTo(node.VisioShape, PinX(node.VisioShape)+moveH, PinY(node.VisioShape)+moveV);
                    newBound.Expand(node.VisioShape);
                }
            }

            return newBound;
        }

        bool IsDrawAtTail(List<GNode> nodes, GrowthDirection direct)
        {
            int nOut = 0, nIn = 0;
            foreach (GNode from in nodes)
            {
                foreach (GNode to in drawnList)
                {
                    if(from.OutputNodes().Contains(to)) nOut++;
                    if(from.InputNodes().Contains(to)) nIn++;
                }
            }

            if (nOut == nIn) return direct.Positive;
            else return nIn > nOut;
        }

        VBoundary NodesBoundary(List<GNode> nodes)
        {
            VBoundary boundary = new(true);
            foreach (var node in nodes)
            {
                if (node.VisioShape != null)
                    boundary.Expand(node.VisioShape);
            }

            return boundary;
        }

        public static double ShapeSheetIU(GNode node, string propName)
        {
            if (node.VisioShape == null) return 0;
            return ShapeSheetIU(node.VisioShape, propName);
        }

        public GNode DropSubgraphBorder(GSubgraph gSubgraph)
        {
            if (gSubgraph.Parent == null) throw new SynException("expected parent of subgraph");

            GNode node = gSubgraph.BorderNode; 
            VBoundary bnd = SubgraphBoundary(gSubgraph);
            Shape shape = CreateShape(node);
            shape.CellsU["Width"].FormulaU = (bnd.Width + GNode.SPACE * 2).ToString();
            shape.CellsU["Height"].FormulaU = (bnd.Height + GNode.SPACE * 2).ToString();
            shape.CellsU["PinX"].FormulaU = bnd.PinX.ToString();
            shape.CellsU["PinY"].FormulaU = bnd.PinY.ToString();
            shape.CellsU["FillPattern"].FormulaU = "0";
            shape.CellsU["VerticalAlign"].FormulaU = "0";
            shape.Text = gSubgraph.Label;
            SetFillForegnd(shape, "config.themeVariables.secondaryColor");
            SetLineColor(shape, "config.themeVariables.secondaryBorderColor");
            SetTextColor(shape, "config.themeVariables.secondaryTextColor");
            shape.CellsU["LineWeight"].FormulaU = "0.75 pt";
            visioApp.DoCmd((short)VisUICmds.visCmdObjectSendToBack);
            gSubgraph.VisioShape = shape;

            return node;
        }

        VBoundary SubgraphBoundary(GSubgraph gSubgraph)
        {
            VBoundary boundary = new VBoundary(true);
            foreach (INode node in gSubgraph.AlignGroupedNodes()) 
            {
                if (node.Container == gSubgraph && node.VisioShape != null) 
                    boundary.Expand(node.VisioShape);
            }
            if (gSubgraph.VisioShape != null) 
                boundary.Expand(gSubgraph.VisioShape);

            foreach (GSubgraph sub in gSubgraph.Subgraphs)
            {
                VBoundary subBoundary = SubgraphBoundary(sub);
                boundary.Expand(subBoundary);
            }

            return boundary;
        }

        public Shape CreateShape(GNode node)
        {
            Shape shape = visioPage.Drop(GetMaster($"{node.ShapeStart}{node.ShapeClose}"), 0, 0);
            shape.Text = node.Label;
            AdjustSize(shape);
            node.VisioShape = shape; 
            return shape;
        }

        public Shape CreateEdge(GEdge edge)
        {
            Master? master = GetMaster("-");
            Shape shape = visioPage.Drop(master, 0, 0);
            SetupEdgeShape(edge, shape);

            return shape;
        }

        void SetupEdgeShape(GEdge edge, Shape shape)
        {
            shape.Text = edge.Text;
            // line type
            switch (edge.LineType)
            {
                case "-": shape.CellsU["LineWeight"].FormulaU = "=0.75 pt"; break;
                case "=": shape.CellsU["LineWeight"].FormulaU = "=0.75 pt"; break;
                case ".": shape.CellsU["LinePattern"].FormulaU = "=2"; break;
                case "~": shape.CellsU["LinePattern"].FormulaU = "=0"; break;
                default: shape.CellsU["LineWeight"].FormulaU = "=0.75 pt"; break;
            }

            // start tag
            // x/o/-/<
            switch (edge.StartTag)
            {
                case "x": shape.CellsU["BeginArrow"].FormulaU = "=24"; break;
                case "o": shape.CellsU["BeginArrow"].FormulaU = "=10"; break;
                case "-": shape.CellsU["BeginArrow"].FormulaU = "=0"; break;
                case "<": shape.CellsU["BeginArrow"].FormulaU = "=4"; break;
                default: shape.CellsU["BeginArrow"].FormulaU = "=0"; break;
            }

            // end tag
            // x/o/-/>
            switch (edge.EndTag)
            {
                case "x": shape.CellsU["EndArrow"].FormulaU = "=24"; break;
                case "o": shape.CellsU["EndArrow"].FormulaU = "=10"; break;
                case "-": shape.CellsU["EndArrow"].FormulaU = "=0"; break;
                case ">": shape.CellsU["EndArrow"].FormulaU = "=4"; break;
                default: shape.CellsU["EndArrow"].FormulaU = "=0"; break;
            }
            
            // 设置连接线颜色
            SetLineColor(shape, "config.themeVariables.defaultLinkColor");
        }


    }
}
