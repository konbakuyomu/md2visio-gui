using md2visio.mermaid.cmn;
using md2visio.struc.figure;
using md2visio.vsdx.@base;

namespace md2visio.main
{
    internal class AppTest
    {
        string testDir = @"C:\Users\Megre\Documents\GitHub\md2visio\md2visio\test";
        string[] files = ["graph", "journey", "packet", "pie", "xy"];
        string outputPath = @"C:\Users\Megre\Desktop";

        public void Test()
        {
            AppConfig config = AppConfig.Instance;
            foreach (string file in files)
            {
                SynContext synContext = new(@$"{testDir}\{file}.md");
                SttMermaidStart.Run(synContext);
                Console.Write(synContext.ToString());

                new FigureBuilderFactory(synContext.NewSttIterator()).Build(outputPath);
            }
            VBuilder.VisioApp.Quit();
        }
    }
}
