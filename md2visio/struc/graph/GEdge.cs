using md2visio.struc.figure;
using System.Text.RegularExpressions;

namespace md2visio.struc.graph
{
    internal class GEdge : Edge
    {
        GNode from = Empty.Get<GNode>(), to = Empty.Get<GNode>();
        string lineType = "-"; // solid/bold/dashed (-/=/.)
        string startTag = "-"; // x/o/-/<
        string endTag = "-"; // x/o/-/>
        public string Text { get; set; } = string.Empty;
        public GNode From { get { return (GNode) from.Container.NodeDict[from.ID]; } }
        public GNode To { get { return (GNode) to.Container.NodeDict[to.ID]; ; } }
        public string LineType
        {
            get { return lineType; }
            set
            {
                if (value.Contains("=")) lineType = "=";
                else if (value.Contains(".")) lineType = ".";
                else if (value.Contains("~")) lineType = "~";
                else lineType = "-";
            }
        }
        public string StartTag
        {
            get { return startTag; }
            set
            {
                startTag = value.Substring(0, 1);
                LineType = value;
            }
        }
        public string EndTag
        {
            get { return endTag; }
            set
            {
                endTag = $"{value.Last()}";
            }
        }
        public GEdge() { }
        public GEdge(GNode from)
        {
            this.from = from;
        }
        public void ConnectTo(GNode to)
        {
            if (from.IsEmpty() || to.IsEmpty()) throw new ArgumentException("connection points can't be empty");

            this.to = to;
            from.AddOutEdge(this);
            to.AddInEdge(this);
        }

        public void Connect(GNode from, GNode to)
        {
            this.from = from;
            ConnectTo(to);
        }

        public GEdge Clone()
        {
            GEdge c = new GEdge();
            c.Text = Text;
            c.lineType = lineType;
            c.startTag = startTag;
            c.endTag = endTag;
            c.from = from;
            c.to = to;

            return c;
        }

        static public bool IsEdgeFragment(string fragment)
        {
            return Regex.IsMatch(fragment, @"^(--|==)>?$");
        }

        static public bool IsEdgeStartFragment(string fragment)
        {
            return Regex.IsMatch(fragment, @"^(--|==)$");
        }

        static public string EdgeEndFragmentPattern(string startFragment)
        {
            if (startFragment == "--") return "-->";
            return "==>";
        }

        static public string GetLineType(string fragment)
        {
            return fragment.Replace(">", "");
        }

        public override string ToString()
        {
            return string.Format("{0}{1}{2}", LineType, Text, Text == string.Empty ? ">" : $"{LineType}>");
        }

        public override bool Equals(object? obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (obj == null || GetType() != obj.GetType()) return false;

            return from.Equals(((GEdge)obj).from) &&
                to.Equals(((GEdge)obj).to) &&
                Text.Equals(((GEdge)obj).Text);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                string edge = ToString();
                int hash = 17;
                hash = hash * 23 + (edge != null ? edge.GetHashCode() : 0);
                hash = hash * 23 + (edge != null ? edge.GetHashCode() : 0);
                return hash;
            }
        }
    }
}
