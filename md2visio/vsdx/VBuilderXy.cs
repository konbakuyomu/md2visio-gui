using md2visio.struc.xy;
using md2visio.vsdx.@base;
using System.Drawing;

namespace md2visio.vsdx
{
    internal class VBuilderXy: VFigureBuilder<XyChart>
    {
        double width, height; // mm
        bool vertical = true;
        VDrawerXy drawer;

        public VBuilderXy(XyChart figure): base(figure)
        {
            drawer = new VDrawerXy(figure, VisioApp);
        }

        protected override void ExecuteBuild()
        {           
            // config
            figure.Config.GetDouble("config.xyChart.width", out width);
            figure.Config.GetDouble("config.xyChart.height", out height);
            figure.Config.GetString("config.xyChart.chartOrientation", out string ori);
            
            width = VShapeDrawer.Pix2MM() * width;
            height = VShapeDrawer.Pix2MM() * height;
            this.Assert(width > 0 && height > 0, "invalid xyChart width/height");

            vertical = (ori != "horizontal");
            
            // draw
            drawer.ChartWidth = width;
            drawer.ChartHeight = height;
            drawer.Vertical = vertical;
            InitXTicks(drawer.XTicks);
            InitYTicks(drawer.YTicks);
            drawer.Draw();
        }

        void InitXTicks(List<string> ticks)
        {
            XyAxis x = figure.XAxis;
            if (x.Values.Count > 0)
            {
                foreach (string tick in x.Values) 
                    ticks.Add(tick);
            }
            else
            {
                int tickNum = Math.Max(figure.Line.Count, figure.Bar.Count);
                this.Assert(tickNum > 0, "invalid xyChart line/bar");

                float start, end;
                if(x.Range != SizeF.Empty) (start, end) = (x.Range.Width, x.Range.Height);
                else (start, end) = (1, tickNum);

                AddTicks(true, ticks, start, end);
            }
        }

        void InitYTicks(List<string> ticks)
        {
            XyAxis y = figure.YAxis;

            // range
            float start = 0, end = 0;
            if(y.Range != SizeF.Empty) 
                (start, end) = (y.Range.Width, y.Range.Height);
            else
            {
                (start, end) = DetermineRange(figure.Bar);
                (float s, float e) = DetermineRange(figure.Line);
                (start, end) = (Math.Min(start, s), Math.Max(end, e));
            }

            // ticks
            AddTicks(false, ticks, start, end);
        }

        void AddTicks(bool xAxis, List<string> ticks, float start, float end)
        {
            float tickSpacing = TickSpacing(xAxis, start, end);
            double tickStep = TickStep(xAxis, start, end, tickSpacing);
            int nDecimal = DecimalNum(end);
            double stop = Math.Max(start, end), tick;
            for (tick = tickStep; tick <= stop; tick += tickStep)
            {
                ticks.Add(ToFixed(tick, nDecimal));
            }
            if (tick != stop) ticks.Add(ToFixed(tick, nDecimal));
            if (start > end) ticks.Reverse();
        }

        float TickSpacing(bool xAxis, float start, float end)
        {
            SizeF textSize = TextSize(Math.Max(start, end));
            return (vertical && xAxis) || (!vertical && !xAxis) ? (textSize.Width + 1) : textSize.Height * 3;
        }

        SizeF TextSize(double num)
        {
            return drawer.MeasureTextSizeMM(TrimZeroEnd(num.ToString()));
        }

        int DecimalNum(double num)
        {
            string n = TrimZeroEnd($"{num}");
            int index = n.LastIndexOf(".");
            if (index == -1) return 0;

            return n.Length - index - 1;
        }

        string TrimZeroEnd(string num)
        {
            return num.TrimEnd(['.', '0']);
        }

        string ToFixed(double num, int pointNum)
        {
            return TrimZeroEnd($"{Math.Round(num, pointNum)}");
        }

        double TickStep(bool xAxis, double start, double end, double tickSpacing)
        {
            this.Assert(start != end, "invalid axis range");
            this.Assert(tickSpacing > 0, "invalid text size");

            double axisStep = Math.Abs(end - start) / ((xAxis ? width : height) / tickSpacing);
            int exponent = (int)Math.Ceiling(Math.Log10(axisStep));
            return Math.Pow(10, exponent);
        }

        (float start, float end) DetermineRange(IEnumerable<object> arr)
        {
            float start = 0, end = 0;
            foreach (object value in figure.Bar)
            {
                if (value == figure.Bar.First())
                {
                    float.TryParse(value.ToString(), out start);
                    end = start;
                }
                else
                {
                    float.TryParse(value.ToString(), out float f);
                    start = Math.Min(start, f);
                    end = Math.Max(end, f);
                }
            }
            return (start, end);
        }

        
   
    }
}
