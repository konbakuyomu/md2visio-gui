using md2visio.mermaid.cmn;
using System.Reflection;
using md2visio.vsdx.@base;

namespace md2visio.struc.figure
{
    internal class FigureBuilderFactory
    {
        string outputFile;
        string? dir = string.Empty, name = string.Empty;
        Dictionary<string, Type> builderDict = TypeMap.BuilderMap;
        SttIterator iter;
        int count = 1;
        bool isFileMode = false; // 标记是否为文件模式

        public FigureBuilderFactory(SttIterator iter)
        {
            this.iter = iter;
            outputFile = iter.Context.InputFile;
        }

        public void Build(string outputFile)
        {
            this.outputFile = outputFile;
            InitOutputPath();
            BuildFigures();
        }
        public void BuildFigures()
        {
            while (iter.HasNext())
            {
                List<SynState> list = iter.Context.StateList;
                for (int pos = iter.Pos + 1; pos < list.Count; ++pos)
                {
                    string word = list[pos].Fragment;
                    if (SttFigureType.IsFigure(word)) BuildFigure(word);
                }
            }            
        }
        public void Quit()
        {
            if (VBuilder.VisioApp != null)
            {
                try
                {
                    // 如果不是显示模式，才退出Visio应用程序
                    if (!md2visio.main.AppConfig.Instance.Visible)
        {
            VBuilder.VisioApp.Quit();
                        VBuilder.VisioApp = null; // 清空静态引用
                    }
                    // 显示模式下保持Visio打开，让用户查看结果
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // 忽略COM异常，可能Visio已经关闭
                    VBuilder.VisioApp = null; // 清空静态引用
                }
            }
        }

        void BuildFigure(string figureType)
        {
            if (!builderDict.ContainsKey(figureType)) 
                throw new NotImplementedException($"'{figureType}' builder not implemented");

            Type type = builderDict[figureType];
            object? obj = Activator.CreateInstance(type, iter);
            MethodInfo? method = type.GetMethod("Build", BindingFlags.Public | BindingFlags.Instance, null,
                new Type[] { typeof(string) }, null);

            // 根据模式选择文件命名策略
            string outputFilePath;
            if (isFileMode)
            {
                // 文件模式：使用用户指定的文件名
                outputFilePath = $"{dir}\\{name}.vsdx";
            }
            else
            {
                // 目录模式：使用计数器区分多个图表
                outputFilePath = $"{dir}\\{name}{count++}.vsdx";
            }

            method?.Invoke(obj, new object[] { outputFilePath });
        }

        void InitOutputPath()
        {
            if (outputFile.ToLower().EndsWith(".vsdx") || File.Exists(outputFile)) // file
            {
                isFileMode = true;
                name = Path.GetFileNameWithoutExtension(outputFile);
                dir = Path.GetDirectoryName(outputFile);
            }
            else // directory
            {
                isFileMode = false;
                if (!Directory.Exists(outputFile))
                    throw new FileNotFoundException($"output path doesn't exist: '{outputFile}'");
                name = Path.GetFileNameWithoutExtension(iter.Context.InputFile);
                dir = Path.GetFullPath(outputFile).TrimEnd(new char[] { '/', '\\' });
            }
        }
    }
}
