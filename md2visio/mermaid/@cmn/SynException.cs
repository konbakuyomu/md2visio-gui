using System.Text;

namespace md2visio.mermaid.cmn
{
    internal class SynException : Exception
    {
        public SynException(string message) : base(message) { }

        public SynException(string message, SynContext context) :
            base(ContextMessage(message, context))
        {

        }

        public SynException(string message, SttIterator iter) :
            base(ContextMessage(message, iter))
        {

        }

        static string ContextMessage(string message, SttIterator iter)
        {
            return $"{message} near {iter.PeekPre().ToString()} {iter.PeekNext().ToString()}";
        }

        static string ContextMessage(string message, SynContext context)
        {
            return $"{message} at line {context.LineNum} near {CurrentLine(context)}";
        }

        static string CurrentLine(SynContext context)
        {
            StringBuilder sb = new StringBuilder();
            for(int i = context.Consumed.Length-1; i>=0; --i)
            {
                char c = context.Consumed[i];
                if (c == '\n') break;
                sb.Insert(0, c);
            }
            int column = sb.Length;
            for (int i = 0; i < context.Incoming.Length; ++i)
            {
                char c = context.Incoming[i];
                if (c == '\n') break;
                sb.Append(c);
            }
            return $"'{sb.ToString()}' (column {column})";
        }
    }
}
