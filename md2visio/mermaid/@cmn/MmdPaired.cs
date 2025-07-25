using System.Text;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.cmn
{
    internal class MmdPaired
    {
        static GroupCollection testGroups = Regex.Match("", "").Groups;
        static public GroupCollection Groups { get { return testGroups; } }

        public static bool IsPaired(string pairStart, string text)
        {
            return IsPaired(pairStart, new StringBuilder(text));
        }

        public static bool IsPaired(string pairStart, StringBuilder textBuilder)
        {
            string pairClose = PairClose(pairStart);
            StringBuilder sb = new StringBuilder();
            bool withinQuote = false;
            int stackCount = 0;
            bool metStart = false;
            for (int i = 0; i < textBuilder.Length; ++i)
            {
                char c = textBuilder[i];
                sb.Append(c);
                if (c == '"') withinQuote = !withinQuote;
                if (withinQuote) continue;

                // test start tag
                int len = Math.Min(pairStart.Length, sb.Length);
                if (sb.ToString(sb.Length - len, len) == pairStart)
                {
                    metStart = true;
                    stackCount++;
                }

                // test close tag
                len = Math.Min(pairClose.Length, sb.Length);
                if (sb.ToString(sb.Length - len, len) == pairClose) stackCount--;

                if (!metStart) continue;
                if (stackCount != 0) continue;

                int Lm = sb.Length - 2 * pairClose.Length;
                int Lc = pairClose.Length;
                testGroups = Regex.Match(textBuilder.ToString(0, sb.Length),
                    @$"^(?s)(?<start>.{{{Lc}}})(?<mid>.{{{Lm}}})(?<close>.{{{Lc}}})").Groups;
                return true;
            }

            return false;
        }

        public static string PairClose(string pairStart)
        {
            // 特殊处理复合配对符号
            if (pairStart == "([") return "])";
            if (pairStart == "[(") return ")]";
            
            // 原有逻辑保持不变（单字符配对）
            StringBuilder sb = new StringBuilder();
            foreach (char c in pairStart)
            {
                sb.Append(PairClose(c));
            }
            return sb.ToString();
        }

        public static char PairClose(char pairStart)
        {
            if (pairStart == '[') return ']';
            else if (pairStart == '{') return '}';
            else if (pairStart == '(') return ')';
            else if (pairStart == '>') return ']';
            else if (pairStart == '\'') return '\'';
            else if (pairStart == '"') return '"';
            else if (pairStart == '`') return '`';

            throw new ArgumentException($"{pairStart}");
        }

        public static string PairPattern(string str)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in str)
            {
                if (c == '{' || c == '}'
                    || c == '(' || c == ')'
                    || c == '[' || c == ']')
                    sb.Append("\\");
                sb.Append(c);
            }
            return sb.ToString();
        }
    }
}
