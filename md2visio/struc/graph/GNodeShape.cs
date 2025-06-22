using md2visio.struc.figure;
using System.Text.RegularExpressions;

namespace md2visio.struc.graph
{
    internal class GNodeShape 
    {
        readonly MmdJsonObj data = new MmdJsonObj();
        GLabel label = new GLabel(string.Empty);

        public GNodeShape()
        {
            data["shape"] = "rect";
            data["label"] = label.Content;
        }

        public string Shape { 
            get => $"{data["shape"]}"; 
            set => data["shape"] = value; 
        }
        public string Label
        {
            get => label.Content = $"{data["label"]}";
            set
            {
                if (value?.Length > 0) 
                    data["label"] = value;
            }
        }

        public string Start { get => ShapeStart(Shape); }

        public string Close { get => ShapeClose(Shape); }

        public string GetData(string name)
        {
            return $"{data[name]}";
        }

        public void SetData(string name, string value)
        {
            data[name] = value;
        }

        public bool HasData(string name)
        {
            return data.HasKey(name);
        }

        public static GNodeShape CreatePaired(string start, string mid, string close)
        {
            GNodeShape shape = new GNodeShape();
            shape.Shape = ShapeName(start, close);
            shape.Label = mid;
            return shape;
        }

        public static GNodeShape CreateExtend(string text)
        {
            GNodeShape shape = new GNodeShape();
            shape.data.UpdateWith(new MmdJsonObj(text));
            return shape;
        }

        public static string ShapeName(string shapeStart, string shapeClose = @"/]")
        {
            switch (shapeStart)
            {
                case ">": return "odd";
                case "[": return "rect";
                case "[[": return "subproc";
                case "{": return "diamond";
                case "{{": return "hex";
                case "(": return "rounded";
                case "((": return "circle";
                case "(((": return "dbl-circ";
                case "[(": return "cyl";
                case @"[\": return shapeClose == @"\]" ? "lean-l" : "trap-t";
                case "[/": return shapeClose == "/]" ? "lean-r" : "trap-b";
                case "([": return "stadium";
            }
            throw new ArgumentException($"unknown shape start '{shapeStart}'");
        }

        public static string ShapeStart(string shapeName)
        {
            switch (shapeName)
            {
                case "odd": return ">";
                case "rect":
                case "rectangle": return "[";
                case "subproc": return "[[";
                case "diamond": return "{";
                case "hex": return "{{";
                case "rounded": return "(";
                case "circle": return "((";
                case "dbl-circ": return "(((";
                case "cyl": return "[(";
                case "lean-l":
                case "trap-t": return @"[\";
                case "lean-r":
                case "trap-b": return "[/";
                case "stadium": return "([";
            }
            throw new ArgumentException($"unknown shape name '{shapeName}'");
        }

        public static string ShapeClose(string shapeName)
        {
            switch (shapeName)
            {
                case "odd":
                case "rect":
                case "rectangle": return "]";
                case "subproc": return "]]";
                case "diamond": return "}";
                case "hex": return "}}";
                case "rounded": return ")";
                case "circle": return "))";
                case "dbl-circ": return ")))";
                case "cyl": return ")]";
                case "lean-l": return @"\]";
                case "trap-t": return "/]";
                case "lean-r": return "/]";
                case "trap-b": return @"\]";
                case "stadium": return "])";
            }
            throw new ArgumentException($"unknown shape name '{shapeName}'");
        }

        static public bool IsShapeFragment(string fragment)
        {
            return IsShapeStartFragment(fragment) ||
                IsShapeCloseFragment(fragment);
        }

        static Regex regShapeStart =
            new(@"^(>|\[{1,2}|\{{1,2}|\({1,3}|\[\(|\[\\|\[/|\(\[)$", RegexOptions.Compiled);
        static public bool IsShapeStartFragment(string fragment)
        {
            return regShapeStart.IsMatch(fragment);
        }

        static Regex regShapeClose =
            new(@"^(\]{1,2}|\}{1,2}|\){1,3}|\)\]|\\\]|/\]|\]\))$", RegexOptions.Compiled);
        static public bool IsShapeCloseFragment(string fragment)
        {
            return regShapeClose.IsMatch(fragment);
        }

        static public string ShapeCloseFragmentPattern(string startFragment)
        {
            switch (startFragment)
            {
                case "[[": return @"\]\]";
                case "[": return @"\]";
                case "{{": return @"\}\}";
                case "{": return @"\}";
                case "(((": return @"\)\)\)";
                case "((": return @"\)\)";
                case "(": return @"\)";
                case ">": return @"\]";
                case "[(": return @"\)\]";
                case @"[\":
                case @"[/": return @"[\\/]]";
                case "([": return @"\]\)";
            }
            return "]";
        }

    }
}
