using md2visio.struc.figure;
using md2visio.vsdx;

namespace md2visio.struc.journey
{
    internal class Journey : Figure
    {
        public Journey() { }

        public HashSet<string> JoinerSet()
        {
            HashSet<string> set = new HashSet<string>();
            foreach (INode section in innerNodes.Values)
            {
                foreach (string joiner in ((JoSection)section).JoinerSet())
                    set.Add(joiner);
            }
            return set;
        }

        public override void ToVisio(string path)
        {
            new VBuilderJo(this).Build(path);
        }
    }
}
