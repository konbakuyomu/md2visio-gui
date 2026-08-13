using md2visio.mermaid.cmn;
using md2visio.Api;
using md2visio.Localization;
using md2visio.vsdx.@base;
using System.Reflection;

namespace md2visio.struc.figure
{
    internal class FigureBuilderFactory
    {
        string outputFile;
        string? dir = string.Empty, name = string.Empty;
        Dictionary<string, Type> builderDict = TypeMap.BuilderMap;
        SttIterator iter;
        int count = 1;
        bool isFileMode = false;
        int figuresBuilt = 0;
        List<string> unsupportedTypes = new();

        // 注入的依赖
        private readonly ConversionContext _context;
        private readonly IVisioSession _session;

        /// <summary>
        /// 构建的图表数量
        /// </summary>
        public int FiguresBuilt => figuresBuilt;

        /// <summary>
        /// 遇到的不支持类型
        /// </summary>
        public IReadOnlyList<string> UnsupportedTypes => unsupportedTypes;

        public FigureBuilderFactory(SttIterator iter, ConversionContext context, IVisioSession session)
        {
            this.iter = iter;
            this._context = context;
            this._session = session;
            outputFile = iter.Context.InputFile;
        }

        public void Build(string outputFile)
        {
            this.outputFile = outputFile;
            // Reset diagnostics for fresh build
            figuresBuilt = 0;
            unsupportedTypes.Clear();

            InitOutputPath();
            BuildFigures();
        }

        /// <summary>
        /// Get supported types string from BuilderMap
        /// </summary>
        static string GetSupportedTypesString()
        {
            return string.Join(", ", TypeMap.BuilderMap.Keys.Distinct().OrderBy(k => k));
        }

        public void BuildFigures()
        {
            if (_context.Debug)
            {
                _context.Log($"[DEBUG] BuildFigures: 开始构建图表");
                _context.Log($"[DEBUG] BuildFigures: iter.HasNext() = {iter.HasNext()}");
                if (iter.Context?.StateList != null)
                {
                    _context.Log($"[DEBUG] BuildFigures: StateList.Count = {iter.Context.StateList.Count}");
                    _context.Log($"[DEBUG] BuildFigures: iter.Pos = {iter.Pos}");

                    for (int i = 0; i < iter.Context.StateList.Count; i++)
                    {
                        var state = iter.Context.StateList[i];
                        _context.Log($"[DEBUG] StateList[{i}]: Type={state.GetType().Name}, Fragment='{state.Fragment}'");
                    }
                }
            }

            // 检查是否找到任何 mermaid 块
            if (iter.Context?.StateList == null || iter.Context.StateList.Count == 0)
            {
                _context.Log(CoreStrings.Get("NoMermaidBlocks"));
                _context.Log(CoreStrings.Get("MermaidBlockHint"));
                return;
            }

            while (iter.HasNext())
            {
                List<SynState> list = iter.Context.StateList;
                bool foundFigure = false;

                for (int pos = iter.Pos + 1; pos < list.Count; ++pos)
                {
                    string word = list[pos].Fragment;

                    if (_context.Debug)
                    {
                        _context.Log($"[DEBUG] BuildFigures: 检查位置 {pos}, Fragment = '{word}'");
                        _context.Log($"[DEBUG] BuildFigures: SttFigureType.IsFigure('{word}') = {SttFigureType.IsFigure(word)}");
                    }

                    if (SttFigureType.IsFigure(word))
                    {
                        foundFigure = true;

                        // 检查是否实现
                        if (!builderDict.ContainsKey(word))
                        {
                            _context.Log(CoreStrings.Format("UnsupportedSkipped", word));
                            if (!unsupportedTypes.Contains(word))
                                unsupportedTypes.Add(word);

                            // 跳过不支持的图表：推进迭代器到下一个 SttMermaidStart 或结束
                            SkipUnsupportedFigure(pos);
                            break;
                        }

                        if (_context.Debug)
                        {
                            _context.Log($"[DEBUG] BuildFigures: 找到图表类型 '{word}'，开始构建");
                        }
                        BuildFigure(word);
                        figuresBuilt++;
                        break;  // BuildFigure 会推进迭代器
                    }
                }

                // 如果遍历完整个列表都没找到图表类型，退出循环
                if (!foundFigure)
                    break;
            }

            // 汇总信息
            if (figuresBuilt == 0)
            {
                _context.Log(CoreStrings.Get("NoDiagramsBuilt"));
                if (unsupportedTypes.Count > 0)
                {
                    _context.Log(CoreStrings.Format("UnsupportedFound", unsupportedTypes.Count, string.Join(", ", unsupportedTypes)));
                    _context.Log(CoreStrings.Format("SupportedTypes", GetSupportedTypesString()));
                }
            }
            else
            {
                _context.Log(CoreStrings.Format("DiagramsBuilt", figuresBuilt));
            }
        }

        /// <summary>
        /// 跳过不支持的图表类型，推进迭代器到下一个图表边界
        /// </summary>
        void SkipUnsupportedFigure(int startPos)
        {
            List<SynState> list = iter.Context.StateList;

            // 从当前位置推进迭代器，直到遇到 SttMermaidClose 或超出范围
            while (iter.HasNext())
            {
                var state = iter.Next();
                if (state.GetType().Name == "SttMermaidClose")
                    break;
            }

            if (_context.Debug)
            {
                _context.Log($"[DEBUG] SkipUnsupportedFigure: 跳过完成，iter.Pos = {iter.Pos}");
            }
        }

        public void Quit()
        {
            // Quit 逻辑已移至 VisioSession.Dispose()
            // 此方法保留为空，供向后兼容
        }

        void BuildFigure(string figureType)
        {
            if (_context.Debug)
            {
                _context.Log($"[DEBUG] BuildFigure: 开始构建图表类型 '{figureType}'");
                _context.Log($"[DEBUG] BuildFigure: builderDict.ContainsKey('{figureType}') = {builderDict.ContainsKey(figureType)}");
            }

            if (!builderDict.ContainsKey(figureType))
                throw new NotImplementedException($"'{figureType}' builder not implemented");

            Type type = builderDict[figureType];

            if (_context.Debug)
            {
                _context.Log($"[DEBUG] BuildFigure: Builder类型 = {type.Name}");
            }

            // 使用注入的 session 和 context 创建 Builder
            object? obj = Activator.CreateInstance(type, iter, _context, _session);
            MethodInfo? method = type.GetMethod("Build", BindingFlags.Public | BindingFlags.Instance, null,
                new Type[] { typeof(string) }, null);

            if (_context.Debug)
            {
                _context.Log($"[DEBUG] BuildFigure: 创建Builder实例成功 = {obj != null}");
                _context.Log($"[DEBUG] BuildFigure: 找到Build方法 = {method != null}");
            }

            string outputFilePath;
            if (isFileMode && count == 1)
            {
                // First figure in file mode: use exact filename
                outputFilePath = $"{dir}\\{name}.vsdx";
                count++;
            }
            else
            {
                // Multiple figures or directory mode: use numbered filenames
                outputFilePath = $"{dir}\\{name}{count++}.vsdx";
            }

            if (_context.Debug)
            {
                _context.Log($"[DEBUG] 构建图表: {figureType}");
                _context.Log($"[DEBUG] 输出模式: {(isFileMode ? "文件模式" : "目录模式")}");
                _context.Log($"[DEBUG] 输出路径: {outputFilePath}");
                _context.Log($"[DEBUG] 输出目录: {dir}");
                _context.Log($"[DEBUG] 文件名: {name}");
            }

            if (_context.Debug)
            {
                _context.Log($"[DEBUG] BuildFigure: 准备调用 {type.Name}.Build('{outputFilePath}')");
            }

            try
            {
                method?.Invoke(obj, new object[] { outputFilePath });

                if (_context.Debug)
                {
                    _context.Log($"[DEBUG] BuildFigure: {type.Name}.Build() 调用完成");
                }
            }
            catch (Exception ex)
            {
                if (_context.Debug)
                {
                    _context.Log($"[DEBUG] BuildFigure: {type.Name}.Build() 调用失败: {ex.Message}");
                    _context.Log($"[DEBUG] BuildFigure: 异常类型: {ex.GetType().Name}");
                    if (ex.InnerException != null)
                    {
                        _context.Log($"[DEBUG] BuildFigure: 内部异常: {ex.InnerException.Message}");
                    }
                }
                throw;
            }

            if (_context.Debug)
            {
                if (File.Exists(outputFilePath))
                {
                    _context.Log($"[DEBUG] ✅ 文件生成成功: {outputFilePath}");
                }
                else
                {
                    _context.Log($"[DEBUG] ❌ 文件生成失败: {outputFilePath}");
                }
            }
        }

        void InitOutputPath()
        {
            if (outputFile.ToLower().EndsWith(".vsdx"))
            {
                isFileMode = true;
                name = Path.GetFileNameWithoutExtension(outputFile);
                dir = Path.GetDirectoryName(outputFile);

                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }
            else if (Directory.Exists(outputFile))
            {
                isFileMode = false;
                name = Path.GetFileNameWithoutExtension(iter.Context.InputFile);
                dir = Path.GetFullPath(outputFile).TrimEnd(new char[] { '/', '\\' });
            }
            else
            {
                throw new ArgumentException(CoreStrings.Format("InvalidOutputPath", outputFile));
            }
        }
    }
}
