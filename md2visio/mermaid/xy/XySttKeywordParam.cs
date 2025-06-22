using md2visio.mermaid.cmn;
using System.Text.RegularExpressions;

namespace md2visio.mermaid.xy
{
    internal class XySttKeywordParam : SttKeywordParam
    {
        public override SynState NextState()
        {
            string keyword = Ctx.StateList.Last().Fragment;
            string param = ExpectedGroups["param"].Value;
            ParseParam(keyword, param);

            return Save(param).Forward<XySttChar>();
        }

        void ParseParam(string keyword, string param)
        {
            if(keyword == "x-axis") ParseXAxis(param);
            else if(keyword == "y-axis") ParseYAxis(param);
        }

        void ParseXAxis(string param)
        {
            Match match = Regex.Match(param,
                "^(\"?(?<title>[^\\[\\n\"]+)\"?)?[^\\S\\n]*((?<vals>\\[[^]]+\\])|(?<range>(?<s>\\d+)[^\\S\\n]*-->[^\\S\\n]*(?<e>\\d+)))?[^\\S\\n]*(?=\\n|$)");
            if (!match.Success) throw new SynException("expected x-axis parameters", Ctx);

            if (match.Groups["title"].Success) AddCompo("title", match.Groups["title"].Value);
            if (match.Groups["vals"].Success)  AddCompo("vals",  match.Groups["vals"].Value);
            if (match.Groups["range"].Success) {
                AddCompo("rangeStart", match.Groups["s"].Value);
                AddCompo("rangeEnd",   match.Groups["e"].Value);
            }
        }

        void ParseYAxis(string param)
        {
            Match match = Regex.Match(param,
                "^(\"?(?<title>[^\\[\\n\"]+)\"?)?[^\\S\\n]*(?<range>(?<s>\\d+)[^\\S\\n]*-->[^\\S\\n]*(?<e>\\d+))?[^\\S\\n]*(?=\\n|$)");
            if (!match.Success) throw new SynException("expected y-axis parameters", Ctx);

            if (match.Groups["title"].Success) AddCompo("title", match.Groups["title"].Value);
            if (match.Groups["range"].Success) { 
                AddCompo("rangeStart", match.Groups["s"].Value);
                AddCompo("rangeEnd",   match.Groups["e"].Value);
            }
        }
    }
}
