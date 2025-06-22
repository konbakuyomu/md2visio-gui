using md2visio.struc.figure;
using md2visio.struc.journey;
using md2visio.vsdx.@base;
using md2visio.vsdx.@tool;
using Microsoft.Office.Interop.Visio;
using System.Drawing;

namespace md2visio.vsdx
{
    internal class VDrawerJo(Journey jo, Application visioApp) : 
        VFigureDrawer<Journey>(jo, visioApp)
    {
        VBoundary drawBound = Empty.Get<VBoundary>();    
        VBoundary allTaskBound = new VBoundary();
        Dictionary<string, VColor> colorMap = new();
        override public void Draw()
        {
            InitColors(figure);

            DrawTitle(figure.Title);
            DrawJoinerTag(figure);
            foreach (INode section in figure.InnerNodes.Values)
            {
                DrawSection((JoSection)section);
            }

            DrawArrow();
        }

        public void DrawArrow()
        {
            Shape shapeArrow = visioPage.Drop(GetMaster("line"), 0, 0);
            SetShapeSheet(shapeArrow, "LineWeight", "2 pt");
            SetShapeSheet(shapeArrow, "EndArrow", "13");
            SetShapeSheet(shapeArrow, "BeginX", $"{allTaskBound.Left}");
            SetShapeSheet(shapeArrow, "BeginY", $"{allTaskBound.Bottom - 0.35}");
            SetShapeSheet(shapeArrow, "EndX", $"{allTaskBound.Right}");
            SetShapeSheet(shapeArrow, "EndY", $"{allTaskBound.Bottom - 0.35}");
        }

        public Shape? DrawTitle(string title)
        {
            Master? master = GetMaster("[]");
            if (master == null) return null;

            Shape shapeTitle = visioPage.Drop(master, 0, 0);
            shapeTitle.Text = title;
            shapeTitle.CellsU["LinePattern"].FormulaU = "=0";
            shapeTitle.CellsU["Char.Size"].FormulaU   = "=14 pt";
            shapeTitle.CellsU["Char.Style"].FormulaU  = "=17"; // bold
            AdjustSize(shapeTitle);
            AlignLeft(shapeTitle, 0);
            drawBound = new VShapeBoundary(shapeTitle);
            drawBound.AlignLeftTop(0, drawBound.Bottom - 0.1);
            return shapeTitle;
        }

        public void DrawJoinerTag(Journey jo)
        {
            List<Shape> textShapeList = new List<Shape>();
            double textWidth = 0;
            foreach(string joiner in jo.JoinerSet())
            {
                Shape shapeText = visioPage.Drop(GetMaster("[]"), 0, 0); 
                shapeText.Text = joiner;
                shapeText.CellsU["LinePattern"].FormulaU = "=0";
                AdjustSize(shapeText);
                textWidth = Math.Max(Width(shapeText), textWidth);
                textShapeList.Add(shapeText);
            }

            VBoundary textBound = drawBound.Clone();
            textBound.AlignRight(-0.1);
            foreach (Shape shapeText in textShapeList)
            {
                SetWidth(shapeText, textWidth.ToString());
                AlignRight(shapeText, textBound.Right);
                AlignTop(shapeText, textBound.Top);

                textBound = new VShapeBoundary(shapeText);
                DrawTagPic((VShapeBoundary) textBound);
                textBound.AlignTop(textBound.Bottom - 0.01);
            }

        }

        void DrawTagPic(VShapeBoundary textBound)
        {
            Shape shapeTag = DropJoinerTag(textBound.Shape.Text);
            MoveTo(shapeTag, textBound.PinX, textBound.PinY);
            AlignRight(shapeTag, textBound.Left - 0.05);
        }

        public Shape? DrawSection(JoSection section)
        {
            Master? master = GetMaster("[]");
            if (master == null) return null;

            Shape ?shapeSection = visioPage.Drop(master, 0, 0);
            shapeSection.Text = section.ID;
            section.VisioShape = shapeSection;
            AdjustSize(shapeSection, new SizeF(0, 1.6f));
            AlignLeftTop(shapeSection, drawBound.Left, drawBound.Top);
            SetFillForegnd(shapeSection, colorMap[section.ID]);
            SetTextColor(shapeSection, "config.themeVariables.primaryTextColor");

            drawBound = new VShapeBoundary(shapeSection);
            drawBound.AlignTop(drawBound.Bottom - 0.15);
            VBoundary boundTasks = drawBound.Clone();
            foreach (JoTask task in section.Tasks)
            {
                Shape? shapeTask = DrawTask(task);
                if (shapeTask == null) continue;                
                
                boundTasks.Expand(shapeTask);
                drawBound.AlignLeft(drawBound.Left + Width(shapeTask) + 0.1);
            }
            allTaskBound.Expand(boundTasks);
            SetShapeSheet(shapeSection, "Width", $"={boundTasks.Width}");
            AlignLeft(shapeSection, boundTasks.Left);
            SetRounding(shapeSection, "1 mm");

            drawBound = new VShapeBoundary(shapeSection);
            drawBound.AlignLeft(drawBound.Right + 0.2);

            return shapeSection;
        }

        public Shape? DrawTask(JoTask task)
        {
            Master? master = GetMaster("task");
            if (master == null) return null;

            Shape shapeTask = visioPage.Drop(master, 0, 0);
            shapeTask.Text = task.Name;
            double minWidth = Width(shapeTask), 
                minHeight = Height(shapeTask);
            AdjustSize(shapeTask);
            if (Width(shapeTask) < minWidth) SetWidth(shapeTask, $"{minWidth}");
            if (Height(shapeTask) < minHeight) SetHeight(shapeTask, $"{minHeight}");
            if(task.Section != null) 
                SetFillForegnd(shapeTask, colorMap[task.Section.ID]);
            SetRounding(shapeTask, "0.6 mm");
            SetTextColor(shapeTask, "config.themeVariables.primaryTextColor");

            AlignLeftTop(shapeTask, drawBound.Left, drawBound.Top);
            VBoundary taskBound = new VShapeBoundary(shapeTask);
            DrawTaskScore(task, taskBound);
            DrawTaskJoiner(task, taskBound);
            return shapeTask;
        }
        public Shape? DrawTaskScore(JoTask joTask, VBoundary taskBound)
        {
            int score = joTask.Score;
            Master? master = GetScoreMaster(score);
            if (master == null) return null;

            Shape shapeScore = visioPage.Drop(master, 0, 0);
            MoveTo(shapeScore, taskBound.PinX, taskBound.Bottom - GetScoreY(shapeScore, score));
            return shapeScore;
        }

        public void DrawTaskJoiner(JoTask joTask, VBoundary taskBound)
        {
            double xshift = 2;
            foreach (string joiner in joTask.Joiners)
            {
                Shape joinerShape = DropJoinerTag(joiner);
                double mm2vu = MM2VisioUnit(joinerShape);
                MoveTo(joinerShape, taskBound.Left + mm2vu * 2 + mm2vu * xshift, taskBound.Top);                
                xshift += 2;
            }
        }

        Shape DropJoinerTag(string joiner)
        {
            Shape joinerShape = visioPage.Drop(GetMaster("(())"), 0, 0);
            SetWidth(joinerShape, "3.8 mm");
            SetHeight(joinerShape, "3.8 mm");
            SetFillForegnd(joinerShape, colorMap[joiner]);
            return joinerShape;
        }

        void InitColors(Journey jo)
        {
            List<INode> sectionList = jo.InnerNodes.Values.ToList();
            List<string> joinerList = jo.JoinerSet().ToList();

            List<string> colors = GetStringList("config.themeVariables.fillType");
            this.Assert(colors.Count > 0, "error reading 'config.themeVariables.fillType?'");
            for (int i = 0; i < jo.InnerNodes.Count; i++)
            {
                 colorMap.Add(sectionList[i].ID, VColor.Create(colors[i % colors.Count]));
            }

            colors = GetStringList("config.journey.actorColours.");
            this.Assert(colors.Count > 0, "error reading 'config.journey.actorColours.?'");
            for (int i = 0; i < joinerList.Count; i++)
            {
                 colorMap.Add(joinerList[i], VColor.Create(colors[i % colors.Count]));
            }
        }

        Master? GetScoreMaster(int score)
        {
            if (score == 3) return GetMaster("thinking");
            else if (score < 3) return GetMaster("crying");
            else return GetMaster("smiling");
        }

        double GetScoreY(Shape scoreShape, int score)
        {
            double u = MM2VisioUnit(scoreShape);
            return (30 + (5 - Math.Min(5,score)) * 10) * u;
        }
    }
}
 