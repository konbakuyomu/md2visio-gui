using md2visio.struc.figure;
using md2visio.vsdx.@tool;
using md2visio.main;
using Microsoft.Office.Interop.Visio;

namespace md2visio.vsdx.@base
{
    internal abstract class VFigureDrawer<T> :
        VShapeDrawer where T : Figure
    {
        protected T figure = Empty.Get<T>();
        protected Config config;

        public VFigureDrawer(T figure, Application visioApp) : base(visioApp)
        {
            this.figure = figure;
            config = figure.Config;
        }

        public abstract void Draw();

        public void SetFillForegnd(Shape shape, string configPath)
        {
            if (config.GetString(configPath, out string sColor))
            {
                SetFillForegnd(shape, VColor.Create(sColor));
            }
        }

        public void SetFillForegnd(Shape shape, VColor color)
        {
            SetShapeSheet(shape, "FillPattern", "1");
            SetShapeSheet(shape, "FillForegnd", $"THEMEGUARD({color.RGB()})");
        }

        public void SetLineColor(Shape shape, string configPath)
        {
            if(config.GetString(configPath, out string color)) {
                SetShapeSheet(shape, "LineColor", $"THEMEGUARD({VColor.Create(color).RGB()})");
            }
        }

        public void SetTextColor(Shape shape, string configPath)
        {
            if (config.GetString(configPath, out string color)) {
                shape.CellsU["Char.Color"].FormulaU = $"THEMEGUARD({VColor.Create(color).RGB()})";
            }
        }

        public List<string> GetStringList(string prefix, int maxCount = 13)
        {
            return figure.Config.GetStringList(prefix, maxCount);
        }

        /// <summary>
        /// 在绘制过程中添加延迟，让用户看到绘制过程
        /// </summary>
        protected void PauseForViewing(int milliseconds = 200)
        {
            if (AppConfig.Instance.Visible)
            {
                System.Threading.Thread.Sleep(milliseconds);
            }
        }

        /// <summary>
        /// 确保Visio应用程序在绘制时可见
        /// </summary>
        protected void EnsureVisible()
        {
            if (AppConfig.Instance.Visible && visioApp != null)
            {
                visioApp.Visible = true;
                visioApp.ActiveWindow?.Activate();
            }
        }

    }
}
