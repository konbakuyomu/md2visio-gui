using System.Text.RegularExpressions;

namespace md2visio.vsdx.@tool
{
    internal class VHSLColor: VColor
    {
        public VHSLColor(float H, float S, float L, float a)
        {
            (this.H, this.S, this.L, this.a) = (H, S, L, a);
            (r, g, b) = HSL2RGB(H, S, L);
            Clamp();
        }

        public static new VColor Create(string color)
        {
            (float H, float S, float L, float a) = ParseColor(color);
            return new VHSLColor(H, S, L, a);
        }

        protected static readonly Regex regHSL = new Regex(@"^\s*hsla?\s*\(\s*(?<h>(?:-\s*)?\d*(.\d*)?)\s*,\s*(?<s>\d*(.\d*)?\s*%?)\s*,\s*(?<l>\d*(.\d*)?\s*%?)\s*(?<a>\d*(.\d*)?)\s*\)\s*$", RegexOptions.Compiled);
        public static bool IsHSL(string color)
        {
            return regHSL.IsMatch(color.Trim().ToLower());
        }

        protected static (float H, float S, float L, float a) ParseColor(string color)
        {
            color = color.Trim().ToLower();
            (float H, float S, float L, float a) = (0, 0, 0, 1);

            Match match;
            if ((match = regHSL.Match(color)).Success)
            {
                float.TryParse(match.Groups["h"].Value, out H);
                TryParsePercent(match.Groups["s"].Value, out S);
                TryParsePercent(match.Groups["l"].Value, out L);
                if (match.Groups["a"].Success)
                    TryParsePercent(match.Groups["a"].Value, out a);
            }            

            return (H, S, L, a);
        }

        static bool TryParsePercent(string percent, out float f)
        {
            if (percent.EndsWith('%'))
            {
                bool success = float.TryParse(percent.TrimEnd('%').Trim(), out float r);
                f = success ? r / 100 : 0;
                return success;
            }
            else
                return float.TryParse(percent.Trim(), out f);
        }
    }
}
