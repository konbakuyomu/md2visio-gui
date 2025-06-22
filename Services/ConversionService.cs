using md2visio.main;
using md2visio.struc.figure;
using System.Diagnostics;

namespace md2visio.GUI.Services
{
    /// <summary>
    /// Mermaid到Visio转换服务
    /// </summary>
    public class ConversionService
    {
        public event EventHandler<ConversionProgressEventArgs>? ProgressChanged;
        public event EventHandler<ConversionLogEventArgs>? LogMessage;

        /// <summary>
        /// 转换MD文件到Visio
        /// </summary>
        /// <param name="inputFile">输入的MD文件路径</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="showVisio">是否显示Visio窗口</param>
        /// <param name="silentOverwrite">是否静默覆盖</param>
        /// <returns>转换结果</returns>
        public async Task<ConversionResult> ConvertAsync(string inputFile, string outputDir, bool showVisio = false, bool silentOverwrite = false)
        {
            return await Task.Run(() => Convert(inputFile, outputDir, showVisio, silentOverwrite));
        }

        /// <summary>
        /// 同步转换方法
        /// </summary>
        private ConversionResult Convert(string inputFile, string outputDir, bool showVisio, bool silentOverwrite)
        {
            try
            {
                ReportProgress(0, "开始转换...");
                ReportLog($"输入文件: {inputFile}");
                ReportLog($"输出目录: {outputDir}");

                // 验证输入文件
                if (!File.Exists(inputFile))
                    return ConversionResult.Error($"输入文件不存在: {inputFile}");

                if (!Path.GetExtension(inputFile).Equals(".md", StringComparison.OrdinalIgnoreCase))
                    return ConversionResult.Error("输入文件必须是 .md 格式");

                // 创建输出目录
                Directory.CreateDirectory(outputDir);
                ReportProgress(20, "准备转换环境...");

                // 构建参数
                var args = new List<string>
                {
                    "/I", $"\"{inputFile}\"",
                    "/O", $"\"{outputDir}\""
                };

                if (showVisio) args.Add("/V");
                if (silentOverwrite) args.Add("/Y");

                ReportProgress(40, "执行转换...");
                ReportLog($"转换参数: {string.Join(" ", args)}");

                // 调用AppConfig进行转换
                var config = new AppConfig();
                if (!config.LoadArguments(args.ToArray()))
                {
                    return ConversionResult.Error("参数解析失败");
                }

                ReportProgress(60, "解析Mermaid内容...");

                // 执行转换
                config.Main();

                ReportProgress(80, "生成Visio文件...");

                // 查找生成的文件
                var outputFiles = Directory.GetFiles(outputDir, "*.vsdx");
                
                ReportProgress(100, "转换完成!");

                if (outputFiles.Length > 0)
                {
                    ReportLog($"成功生成 {outputFiles.Length} 个文件:");
                    foreach (var file in outputFiles)
                    {
                        ReportLog($"  - {Path.GetFileName(file)}");
                    }
                    return ConversionResult.Success(outputFiles);
                }
                else
                {
                    return ConversionResult.Error("转换完成但未找到输出文件");
                }
            }
            catch (NotImplementedException ex)
            {
                ReportLog($"功能未实现: {ex.Message}");
                return ConversionResult.Error($"该图表类型暂未支持: {ex.Message}");
            }
            catch (Exception ex)
            {
                ReportLog($"转换出错: {ex.Message}");
                return ConversionResult.Error($"转换失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检测MD文件中的Mermaid图表类型
        /// </summary>
        /// <param name="filePath">MD文件路径</param>
        /// <returns>检测到的图表类型列表</returns>
        public List<string> DetectMermaidTypes(string filePath)
        {
            var types = new List<string>();
            
            try
            {
                var content = File.ReadAllText(filePath);
                var lines = content.Split('\n');
                
                bool inMermaidBlock = false;
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    
                    if (trimmed.StartsWith("```mermaid"))
                    {
                        inMermaidBlock = true;
                        continue;
                    }
                    
                    if (trimmed.StartsWith("```") && inMermaidBlock)
                    {
                        inMermaidBlock = false;
                        continue;
                    }
                    
                    if (inMermaidBlock && !string.IsNullOrWhiteSpace(trimmed))
                    {
                        // 检测图表类型
                        var words = trimmed.Split(' ');
                        if (words.Length > 0)
                        {
                            var type = words[0].ToLower();
                            if (!types.Contains(type))
                            {
                                types.Add(type);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ReportLog($"检测文件类型时出错: {ex.Message}");
            }
            
            return types;
        }

        private void ReportProgress(int percentage, string message)
        {
            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs(percentage, message));
        }

        private void ReportLog(string message)
        {
            LogMessage?.Invoke(this, new ConversionLogEventArgs(DateTime.Now, message));
        }
    }

    /// <summary>
    /// 转换结果
    /// </summary>
    public class ConversionResult
    {
        public bool IsSuccess { get; set; }
        public string? ErrorMessage { get; set; }
        public string[]? OutputFiles { get; set; }

        public static ConversionResult Success(string[] outputFiles)
        {
            return new ConversionResult { IsSuccess = true, OutputFiles = outputFiles };
        }

        public static ConversionResult Error(string message)
        {
            return new ConversionResult { IsSuccess = false, ErrorMessage = message };
        }
    }

    /// <summary>
    /// 转换进度事件参数
    /// </summary>
    public class ConversionProgressEventArgs : EventArgs
    {
        public int Percentage { get; }
        public string Message { get; }

        public ConversionProgressEventArgs(int percentage, string message)
        {
            Percentage = percentage;
            Message = message;
        }
    }

    /// <summary>
    /// 转换日志事件参数
    /// </summary>
    public class ConversionLogEventArgs : EventArgs
    {
        public DateTime Timestamp { get; }
        public string Message { get; }

        public ConversionLogEventArgs(DateTime timestamp, string message)
        {
            Timestamp = timestamp;
            Message = message;
        }
    }
} 