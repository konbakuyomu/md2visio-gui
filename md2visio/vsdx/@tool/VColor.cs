namespace md2visio.vsdx.@tool
{
    internal abstract class VColor
    {
        protected float r, g, b; // [0, 255]
        protected float a = 1;   // [0, 1]
        protected float H;       // [0, 360]
        protected float S, L;    // [0, 1]

        public static VColor Create(string color) {
            if(VNamedColor.IsNamed(color))  return VNamedColor.Create(color);
            if(VRGBColor.IsRGB(color))      return VRGBColor.Create(color);
            if(VHSLColor.IsHSL(color))      return VHSLColor.Create(color);

            throw new ArgumentException($"Unsupported color format '{color.Trim()}'");
        }

        public string RGB()
        {
            return string.Format("rgb({0:F0}, {1:F0}, {2:F0})",
                Math.Round(r), Math.Round(g), Math.Round(b));
        }

        public string RGBA()
        {
            return string.Format("rgba({0:F0}, {1:F0}, {2:F0}, {3:F0})",
                Math.Round(r), Math.Round(g), Math.Round(b), Math.Round(a));
        }

        public string HexRGB()
        {
            return string.Format("#{0:X2}{1:X2}{2:X2}",
                Math.Round(r), Math.Round(g), Math.Round(b));
        }

        public string HexARGB()
        {
            return string.Format("#{0:X2}{1:X2}{2:X2}{3:X2}",
                Math.Round(a * 255), Math.Round(r), Math.Round(g), Math.Round(b));
        }
        public string HLS()
        {
            float s = S * 100, l = L * 100;
            return $"hsl({H:F0}, {s:F2}%, {l:F2}%)";
        }
        public string HLSA()
        {
            float s = S * 100, l = L * 100;
            return $"hsla({H:F0}, {s:F2}%, {l:F2}%, {a:F2})";
        }               

        protected static (float H, float S, float L) RGB2HSL(float r, float g, float b)
        {
            // r, g, b -> [0, 1]
            (r, g, b) = (r / 255, g / 255, b / 255);

            // 找到最大值和最小值
            float max = Math.Max(r, Math.Max(g, b));
            float min = Math.Min(r, Math.Min(g, b));

            // 计算 L (亮度)
            float l = (max + min) / 2;

            // 计算 S (饱和度)
            float s;
            if (max == min)
            {
                s = 0; // 饱和度为 0，表示灰色
            }
            else
            {
                s = l > 0.5 ? (max - min) / (2 - max - min) : (max - min) / (max + min);
            }

            // 计算 H (色调)
            float h = 0;
            if (max == min)
            {
                h = 0; // 无色调
            }
            else
            {
                if (max == r)
                {
                    h = (g - b) / (max - min);
                }
                else if (max == g)
                {
                    h = 2 + (b - r) / (max - min);
                }
                else if (max == b)
                {
                    h = 4 + (r - g) / (max - min);
                }
                h *= 60; // 转换为角度
                if (h < 0) h += 360; // 确保在 [0, 360] 范围内
            }

            return (h, s, l);
        }
        protected static (float r, float g, float b) HSL2RGB(float H, float S, float L)
        {
            float r, g, b;
            if (S == 0)
            {
                r = g = b = L; // 如果饱和度为 0，则颜色是灰色
            }
            else
            {
                // 根据亮度和区域计算颜色分量
                float HueToRgb(float p, float q, float t)
                {
                    if (t < 0) t += 1;
                    if (t > 1) t -= 1;
                    if (t < 1 / 6f) return p + (q - p) * 6 * t;
                    if (t < 1 / 2f) return q;
                    if (t < 2 / 3f) return p + (q - p) * (2 / 3f - t) * 6;
                    return p;
                }

                float q = L < 0.5f ? L * (1 + S) : L + S - L * S;
                float p = 2 * L - q;

                // 计算 RGB 分量
                r = HueToRgb(p, q, H / 360f + 1 / 3f);
                g = HueToRgb(p, q, H / 360f);
                b = HueToRgb(p, q, H / 360f - 1 / 3f);
            }

            // 将浮点数转换为 [0, 255] 的整数值
            return (r * 255, g * 255, b * 255);
        }
        protected void Clamp()
        {
            (r, g, b, a) = (Math.Clamp(r, 0, 255), Math.Clamp(g, 0, 255), Math.Clamp(b, 0, 255),
                Math.Clamp(a, 0, 1));
            (H, S, L) = (H % 360, Math.Clamp(S, 0, 1), Math.Clamp(L, 0, 1));
        }
    }
}
