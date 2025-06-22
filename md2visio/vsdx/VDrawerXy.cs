using md2visio.struc.xy;
using Microsoft.Office.Interop.Visio;
using md2visio.vsdx.@base;

namespace md2visio.vsdx
{
    internal class VDrawerXy(XyChart figure, Application visioApp) : 
        VFigureDrawer<XyChart>(figure, visioApp)
    {
        public double ChartWidth { get; set; }
        public double ChartHeight { get; set; }
        public bool Vertical { get; set; }
        public List<string> XTicks { get; set; } = [];
        public List<string> YTicks { get; set; } = [];

        public override void Draw()
        {
            

            DrawAxis();
        }

        void DrawAxis()
        {
            Shape x = visioPage.Drop(GetMaster("x-axis"), 0, 0);
            SetShapeSheet(x, "BeginX", "0");
            SetShapeSheet(x, "BeginY", "0");
            SetShapeSheet(x, "EndX", $"{ChartWidth} mm");
            SetShapeSheet(x, "EndY", "0");

            Shape y = visioPage.Drop(GetMaster("y-axis"), 0, 0);
            SetShapeSheet(y, "BeginX", "0");
            SetShapeSheet(y, "BeginY", "0");
            SetShapeSheet(y, "EndX", "0");
            SetShapeSheet(y, "EndY", $"{ChartHeight} mm");
        }
    }
}
