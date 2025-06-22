using md2visio.mermaid.cmn;
using md2visio.mermaid.pie;
using md2visio.struc.figure;

namespace md2visio.struc.pie
{
    internal class PieBuilder(SttIterator iter) : FigureBuilder(iter)
    {
        readonly Pie pie = new();

        public override void Build(string outputFile)
        {
            while (iter.HasNext())
            {
                SynState cur = iter.Next();
                if (cur is SttMermaidStart) { }
                if (cur is SttMermaidClose)
                {
                    pie.ToVisio(outputFile);
                    break;
                }
                if (cur is PieSttKeyword) { BuildKeyword(); }
                if (cur is PieSttTuple) { BuildDataItem(); }
                if (cur is SttComment) { pie.Config.LoadUserDirectiveFromComment(cur.Fragment); }
                if (cur is SttFrontMatter) { pie.Config.LoadUserFrontMatter(cur.Fragment); }
            }
        }

        void BuildKeyword()
        {
            string kw = iter.Current.Fragment;
            SynState next = iter.PeekNext();

            if (kw == "pie") {
                if(next is PieSttKeywordParam)
                {
                    var dict = ((PieSttKeywordParam) iter.Next()).ParsePieParam();
                    pie.ShowData = dict.ContainsKey("showData");
                    if(dict.TryGetValue("title", out string? value)) pie.Title = value;
                }
            }
            else if (kw == "title")
            {
                if (next is not PieSttKeywordParam) throw new SynException("expected title", iter);
                pie.Title = iter.Next().Fragment;
            }
        }

        void BuildDataItem()
        {
            PieSttTuple tuple = (PieSttTuple) iter.Current;
            PieDataItem dataItem = pie.RetrieveNode<PieDataItem>(tuple.CompoList.Get(1));

            if (!double.TryParse(tuple.CompoList.Get(2), out double data) ||
                data <= 0)
                throw new SynException("data must be a positive number", iter);
            dataItem.Data = data;

            pie.AddInnerNode(dataItem);
        }
    }
}
