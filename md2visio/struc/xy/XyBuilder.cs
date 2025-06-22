using md2visio.mermaid.cmn;
using md2visio.mermaid.xy;
using md2visio.struc.figure;

namespace md2visio.struc.xy
{
    internal class XyBuilder(SttIterator iter) : FigureBuilder(iter)
    {
        readonly XyChart xy = new();

        public override void Build(string outputFile)
        {
            while (iter.HasNext())
            {
                SynState cur = iter.Next();
                if (cur is SttMermaidStart) { }
                if (cur is SttMermaidClose) { xy.ToVisio(outputFile);  break; }
                if (cur is XySttKeyword)    { BuildKeyword(); }
                if (cur is SttComment)      { xy.Config.LoadUserDirectiveFromComment(cur.Fragment); }
                if (cur is SttFrontMatter)  { xy.Config.LoadUserFrontMatter(cur.Fragment); }
            }
        }

        void BuildKeyword()
        {
            string kw = iter.Current.Fragment;
            SynState next = iter.PeekNext();

            if (kw == "xychart" || kw == "xychart-beta")
            {
                if (next is XySttKeywordParam) 
                    xy.FrontMatter.SetValue("init.xyChart.chartOrientation", iter.Next().Fragment);
            }
            else if (kw == "title")
            {
                if (next is not XySttKeywordParam) throw new SynException("expected title", iter);
                xy.Title = iter.Next().Fragment;
            }
            else if(kw == "x-axis")
            {
                if (next is not XySttKeywordParam) throw new SynException("expected x-axis parameters", iter);
                CompoDict compo = iter.Next().CompoList;
                xy.XAxis = new XyAxis(compo);
            }
            else if(kw == "y-axis")
            {
                if (next is not XySttKeywordParam) throw new SynException("expected y-axis parameters", iter);
                CompoDict compo = iter.Next().CompoList;
                xy.YAxis = new XyAxis(compo);
            }
            else if(kw == "line")
            {
                if (next is not XySttKeywordParam) throw new SynException("expected line parameters", iter);
                xy.Line.Load(iter.Next().Fragment);
            }
            else if(kw == "bar")
            {
                if (next is not XySttKeywordParam) throw new SynException("expected bar parameters", iter);
                xy.Bar.Load(iter.Next().Fragment);
            }
        }
    }
}
