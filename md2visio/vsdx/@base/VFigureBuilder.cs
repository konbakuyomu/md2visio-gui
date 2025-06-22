using md2visio.struc.figure;

namespace md2visio.vsdx.@base
{
    internal abstract class VFigureBuilder<T> :
        VBuilder where T : Figure
    {
        protected T figure;

        public VFigureBuilder(T figure)
        {
            this.figure = figure;
        }

        public void Build(string outputFile)
        {
            ExecuteBuild();
            VShapeDrawer.RemoveShadow();
            SaveAndClose(outputFile);
        }

        protected abstract void ExecuteBuild();
    }
}
