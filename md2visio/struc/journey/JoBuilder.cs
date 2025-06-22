using md2visio.mermaid.cmn;
using md2visio.mermaid.journey;
using md2visio.struc.figure;

namespace md2visio.struc.journey
{
    internal class JoBuilder : FigureBuilder
    {
        Journey journey = new Journey();
        JoSection curSection = Empty.Get<JoSection>();
        public JoBuilder(SttIterator iter) : base(iter) { }

        public override void Build(string outputFile)
        {
            while (iter.HasNext())
            {
                SynState cur = iter.Next();
                if (cur is SttMermaidStart) { }
                if (cur is SttMermaidClose)
                {
                    journey.ToVisio(outputFile);
                    break;
                }
                if (cur is JoSttKeyword)    { BuildKeyword(); }
                if (cur is JoSttTriple)     { BuildTask(); }
                if (cur is SttComment)      { journey.Config.LoadUserDirectiveFromComment(cur.Fragment); }
                if (cur is SttFrontMatter)  { journey.Config.LoadUserFrontMatter(cur.Fragment); }
            }
        }

        void BuildKeyword()
        {
            string kw = iter.Current.Fragment;
            SynState next = iter.PeekNext();

            if (kw == "journey") { }
            else if (kw == "title")
            {
                if (next is not JoSttKeywordParam) throw new SynException("expected title", iter);
                journey.Title = iter.Next().Fragment;
            }
            else if (kw == "section")
            {
                if (next is not JoSttKeywordParam) throw new SynException("expected section title", iter);

                curSection = journey.RetrieveNode<JoSection>(iter.Next().Fragment);
                journey.AddInnerNode(curSection);
            }

        }

        void BuildTask()
        {
            if (curSection.IsEmpty()) throw new SynException("expected section", iter);

            JoSttTriple triple = (JoSttTriple)iter.Current;
            CompoDict compo = triple.CompoList;
            JoTask task = new JoTask(compo.Get(1), compo.Get(2), compo.Get(3));
            curSection.AddTask(task);
        }
    }
}
