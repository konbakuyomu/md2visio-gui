using md2visio.mermaid.graph;
using System.Text;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.cmn
{
    internal class SynContext
    {
        List<SynState> stateList = new List<SynState>();

        StringBuilder incoming = new StringBuilder();
        StringBuilder consumed = new StringBuilder();
        StringBuilder cache = new StringBuilder();
        GroupCollection expectedGroup = Regex.Match("", "").Groups;

        public int LineNum { get { return GetLineNum(); } }
        public string InputFile { get; set; } = string.Empty;
        public GroupCollection ExpectedGroups { get { return expectedGroup; } }
        public GroupCollection TestGroups { get; set; } = Regex.Match("", "").Groups;
        public List<SynState> StateList { get { return stateList; } }
        public StringBuilder Incoming { get { return incoming; } }
        public StringBuilder Consumed { get { return consumed; } }
        public StringBuilder Cache { get { return cache; } }

        public SynContext() { }        

        public SynContext(string inputFile)
        {
            InputFile = inputFile;

            try {
                string[] lines = File.ReadAllLines(inputFile);

                foreach (string line in lines) incoming.Append(line).Append('\n');
            } catch (Exception ex) { 
                Console.Error.WriteLine(ex.Message);
            }                       
        }

        public string? Peek(int length = 1)
        {
            if (length > incoming.Length || length < -consumed.Length) return null;
            if (length >= 0) return incoming.ToString(0, length);
            return consumed.ToString(consumed.Length + length, -length);
        }
        public string? Take(int length = 1)
        {
            if (length < 0 || length > incoming.Length) return null;

            string take = incoming.ToString(0, length);
            consumed.Append(take);
            cache.Append(take);
            Skip(length);
            return take;
        }
        public void Skip(int length = 1)
        {
            incoming.Remove(0, Math.Min(length, incoming.Length));
        }

        public string? Slide(int length = 1)
        {
            if (length < 0 || length > incoming.Length)
                throw new ArgumentOutOfRangeException("length");

            string slide = incoming.ToString(0, length);
            consumed.Append(slide);
            Skip(length);
            return slide;
        }

        public void Restore(int length)
        {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(length);

            length = Math.Min(length, consumed.Length);

            int start = consumed.Length - length;
            incoming.Insert(0, consumed.ToString(start, length));
            consumed.Remove(start, length);
        }

        public void Restore(string text)
        {
            this.incoming.Insert(0, text);
        }

        public SynState LastState()
        {
            if (stateList.Count < 1) return EmptyState.Instance;

            return stateList.Last();
        }

        public SynState LastNonFinishState()
        {
            for (int i = stateList.Count - 1; i >= 0; --i)
            {
                if (stateList[i] is not SttFinishFlag) return stateList[i];
            }

            return EmptyState.Instance;
        }

        public bool Expect(string pattern, bool multiline = false)
        {
            Match match = Regex.Match(incoming.ToString(), $"^(?<tail>{pattern})",
                multiline ? RegexOptions.Multiline : RegexOptions.None);
            if (match.Success)
            {
                consumed.Append(match.Groups[0].Value);
                expectedGroup = match.Groups;
                incoming.Remove(0, match.Length);
                return true;
            }
            return false;
        }
        public bool Until(string pattern, bool multiline = true)
        {
            Match match = Regex.Match(incoming.ToString(0, incoming.Length), $"^(?<head>.*?)(?<tail>{pattern})",
                multiline ? RegexOptions.Multiline : RegexOptions.None);
            if (match.Success)
            {
                consumed.Append(match.Groups[0].Value);
                expectedGroup = match.Groups;
                incoming.Remove(0, match.Index + match.Length);
                return true;
            }
            return false;
        }

        public bool Test(string pattern)
        {
            Match match = Regex.Match(Incoming.ToString(), pattern);
            if (!match.Success) return false;

            TestGroups = match.Groups;
            return true;
        }

        public void AddState(SynState state)
        {
            if (stateList.Count > 0)
            {
                if (stateList.Last() is SttFinishFlag &&
                    state is SttFinishFlag) return;
            }
            stateList.Add(state);
        }

        public (bool Success, SynState Container) FindContainerType(string stateNamePattern)
        {
            for (int i = stateList.Count - 1; i >= 0; i--)
            {
                SynState state = stateList[i];
                string typeName = state.GetType().Name;
                if (Regex.IsMatch(typeName, $"^({stateNamePattern})$")) return (true, state);
            }
            return (false, EmptyState.Instance);
        }

        public (bool Success, string Fragment) FindContainerFrag(string fragmentPattern)
        {
            for (int i = stateList.Count - 1; i >= 0; i--)
            {
                string frag = stateList.ElementAt(i).Fragment;
                if (Regex.IsMatch(frag, $"^({fragmentPattern})$")) return (true, frag);
            }
            return (false, string.Empty);
        }

        public bool WithinKeyword()
        {
            for (int i = stateList.Count - 1; i >= 0; i--)
            {
                SynState state = stateList.ElementAt(i);
                if (state is SttFinishFlag) return false;
                if (GSttKeyword.IsKeyword(state.Fragment)) return true;
            }
            return false;
        }

        public void ClearCache() { cache.Clear(); }

        public SttIterator NewSttIterator()
        {
            return new SttIterator(this);
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (SynState state in stateList)
            {
                sb.Append(state.ToString()).Append('\n');
            }
            return sb.ToString();
        }

        int GetLineNum()
        {
            int line = 1;
            for (int i = 0; i < consumed.Length; ++i)
                if (consumed[i] == '\n') ++line;
            return line;            
        }
    }
}
