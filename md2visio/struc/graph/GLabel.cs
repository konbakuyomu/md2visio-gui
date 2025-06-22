using System.Text.RegularExpressions;

namespace md2visio.struc.graph
{
    internal class GLabel 
    {
        string content = string.Empty;
        string markdown = string.Empty;

        protected GLabel() { }

        public GLabel(string content)
        {
            Content = content;
        }

        public string Content
        {
            get { return content; }
            set
            {
                if (value?.Length > 0)
                {
                    content = value;
                    ParseQuote();
                    ParseMarkdown();
                }
            }
        }

        private void ParseQuote()
        {
            Match match = Regex.Match(content, @"(^""(?<ct>.+)""$)|(^'(?<ct>.+)')$");
            if (match.Success) content = match.Groups["ct"].Value;
        }

        private void ParseMarkdown()
        {
            Match match = Regex.Match(content, @"""?`([^`]+)`""?");
            if (match.Success) markdown = match.Groups[1].Value;
        }

        public override string ToString()
        {
            if (markdown.Length == 0) return content;
            return markdown;
        }

    }
}
