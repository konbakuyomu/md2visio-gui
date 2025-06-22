using System.Globalization;
using System.Text.RegularExpressions;

namespace md2visio.vsdx.@tool
{
    internal class VRGBColor: VColor
    {
        public VRGBColor(float r, float g, float b): this(r, g, b, 1)
        {
        }
        public VRGBColor(float r, float g, float b, float a)
        {
            (this.r, this.g, this.b, this.a) = (r, g, b, a);
            (H, S, L) = RGB2HSL(r, g, b);
            Clamp();
        }
        public static new VColor Create(string color)
        {
            (float r, float g, float b, float a) = ParseColor(color);
            return new VRGBColor(r, g, b, a);
        }

        public static bool IsRGB(string color)
        {
            color = color.Trim().ToLower();
            return regRGB.IsMatch(color) || 
                regHexRRGGBB.IsMatch(color) || 
                regHexRGB.IsMatch(color);
        }

        protected static readonly Regex regRGB = new Regex(@"^\s*rgba?\s*\(\s*(?<r>\d*(.\d*)?)\s*,\s*(?<g>\d*(.\d*)?)\s*,\s*(?<b>\d*(.\d*)?)\s*(?<a>\d*(.\d*)?)\s*\)\s*$", RegexOptions.Compiled);
        protected static readonly Regex regHexRRGGBB = new Regex(@"^\s*#\s*(?<a>[0-9a-f]{2})?\s*(?<r>[0-9a-f]{2})\s*(?<g>[0-9a-f]{2})\s*(?<b>[0-9a-f]{2})\s*$", RegexOptions.Compiled);
        protected static readonly Regex regHexRGB = new Regex(@"^\s*#\s*(?<r>[0-9a-f])\s*(?<g>[0-9a-f])\s*(?<b>[0-9a-f])\s*$", RegexOptions.Compiled);
        protected static (float r, float g, float b, float a) ParseColor(string color)
        {
            color = color.Trim().ToLower();
            (float r, float g, float b, float a) = (0, 0, 0, 1);

            Match match;
            // RGB
            if ((match = regRGB.Match(color)).Success)
            {
                float.TryParse(match.Groups["r"].Value, out r);
                float.TryParse(match.Groups["g"].Value, out g);
                float.TryParse(match.Groups["b"].Value, out b);
                if (match.Groups["a"].Success)
                    float.TryParse(match.Groups["a"].Value, out a);
            }
            // Hex #RRGGBB
            else if ((match = regHexRRGGBB.Match(color)).Success)
            {
                try { r = int.Parse(match.Groups["r"].Value, NumberStyles.HexNumber); } catch { }
                try { g = int.Parse(match.Groups["g"].Value, NumberStyles.HexNumber); } catch { }
                try { b = int.Parse(match.Groups["b"].Value, NumberStyles.HexNumber); } catch { }
                if (match.Groups["a"].Success)
                    try { a = int.Parse(match.Groups["a"].Value, NumberStyles.HexNumber); } catch { }
            }
            // Hex #RGB            
            else if ((match = regHexRGB.Match(color)).Success)
            {
                try { r = int.Parse(DoubleHex(match.Groups["r"].Value), NumberStyles.HexNumber); } catch { }
                try { g = int.Parse(DoubleHex(match.Groups["g"].Value), NumberStyles.HexNumber); } catch { }
                try { b = int.Parse(DoubleHex(match.Groups["b"].Value), NumberStyles.HexNumber); } catch { }
                if (match.Groups["a"].Success)
                    try { a = int.Parse(match.Groups["a"].Value, NumberStyles.HexNumber); } catch { }
            }

            return (r, g, b, a);
        }

        static string DoubleHex(string value)
        {
            return $"{value}{value}";
        }
    }
}
