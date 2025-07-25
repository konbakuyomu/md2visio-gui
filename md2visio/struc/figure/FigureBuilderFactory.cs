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

            // 添加调试日志
            if (md2visio.main.AppConfig.Instance.Debug)
            {
                Console.WriteLine($"[DEBUG] 构建图表: {figureType}");
                Console.WriteLine($"[DEBUG] 输出模式: {(isFileMode ? "文件模式" : "目录模式")}");
                Console.WriteLine($"[DEBUG] 输出路径: {outputFilePath}");
                Console.WriteLine($"[DEBUG] 输出目录: {dir}");
                Console.WriteLine($"[DEBUG] 文件名: {name}");
            }

            method?.Invoke(obj, new object[] { outputFilePath });

            // 验证文件是否成功生成
            if (md2visio.main.AppConfig.Instance.Debug)
            {
                if (File.Exists(outputFilePath))
                {
                    Console.WriteLine($"[DEBUG] ✅ 文件生成成功: {outputFilePath}");
                }
                else
                {
                    Console.WriteLine($"[DEBUG] ❌ 文件生成失败: {outputFilePath}");
                }
            }
        }

        void InitOutputPath()
        {
            // 优先检查是否指定了具体的 .vsdx 文件路径
            if (outputFile.ToLower().EndsWith(".vsdx"))
            {
                // 文件模式：用户指定了具体的输出文件名
                isFileMode = true;
                name = Path.GetFileNameWithoutExtension(outputFile);
                dir = Path.GetDirectoryName(outputFile);
                
                // 确保输出目录存在
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }
            else if (Directory.Exists(outputFile))
            {
                // 目录模式：用户指定了输出目录
                isFileMode = false;
                name = Path.GetFileNameWithoutExtension(iter.Context.InputFile);
                dir = Path.GetFullPath(outputFile).TrimEnd(new char[] { '/', '\\' });
            }
            else
            {
                // 既不是 .vsdx 文件也不是存在的目录
                throw new ArgumentException($"输出路径无效: '{outputFile}'。请指定一个 .vsdx 文件路径或现有目录。");
            }
        }
    }
}
