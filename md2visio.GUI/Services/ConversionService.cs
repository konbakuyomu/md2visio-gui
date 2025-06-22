using md2visio.main;
using md2visio.struc.figure;
using System.Diagnostics;

namespace md2visio.GUI.Services
{
    /// <summary>
    /// Mermaid到Visio转换服务
    /// </summary>
    public class ConversionService : IDisposable
    {
        public event EventHandler<ConversionProgressEventArgs>? ProgressChanged;
        public event EventHandler<ConversionLogEventArgs>? LogMessage;

        private AppConfig? _appConfig;
        private bool _disposed = false;

        /// <summary>
        /// 转换MD文件到Visio
        /// </summary>
        /// <param name="inputFile">输入的MD文件路径</param>
        /// <param name="outputDir">输出目录</param>
        /// <param name="showVisio">是否显示Visio窗口</param>
        /// <param name="silentOverwrite">是否静默覆盖</param>
        /// <returns>转换结果</returns>
        public async Task<ConversionResult> ConvertAsync(string inputFile, string outputDir, string? fileName = null, bool showVisio = false, bool silentOverwrite = false)
        {
            return await Task.Run(() => Convert(inputFile, outputDir, fileName, showVisio, silentOverwrite));
        }

        /// <summary>
        /// 同步转换方法
        /// </summary>
        private ConversionResult Convert(string inputFile, string outputDir, string? fileName, bool showVisio, bool silentOverwrite)
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

                // 构建输出路径：如果提供了文件名，使用完整路径；否则使用目录
                string outputPath;
                if (!string.IsNullOrEmpty(fileName))
                {
                    // 确保文件名有.vsdx扩展名
                    if (!fileName.EndsWith(".vsdx", StringComparison.OrdinalIgnoreCase))
                        fileName += ".vsdx";
                    outputPath = Path.Combine(outputDir, fileName);
                }
                else
                {
                    outputPath = outputDir;
                }

                // 构建参数 (不加引号，让AppConfig直接处理路径)
                var args = new List<string>
                {
                    "/I", inputFile,
                    "/O", outputPath
                };

                if (showVisio) args.Add("/V");
                if (silentOverwrite) args.Add("/Y");

                ReportProgress(40, "执行转换...");
                ReportLog($"转换参数: {string.Join(" ", args)}");

                // 如果需要显示Visio窗口，则创建一个由服务管理的AppConfig实例
                // 并将其设置为全局静态实例，以供核心库的其他部分使用
                if (showVisio)
                {
                    _appConfig ??= new AppConfig();
                    AppConfig.Instance = _appConfig;
                }
                else
                {
                    // 对于非显示模式，使用临时的默认实例
                    AppConfig.Instance = new AppConfig();
                }

                if (!AppConfig.Instance.LoadArguments(args.ToArray()))
                {
                    return ConversionResult.Error("参数解析失败");
                }

                ReportProgress(60, "解析Mermaid内容...");

                // 执行转换
                AppConfig.Instance.Main();

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
            catch (ApplicationException ex)
            {
                // 这是从VBuilder抛出的Visio相关异常
                ReportLog($"Visio操作失败: {ex.Message}");
                return ConversionResult.Error($"Visio操作失败: {ex.Message}");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ReportLog($"COM异常: {ex.Message} (HRESULT: 0x{ex.HResult:X8})");
                return ConversionResult.Error($"COM组件异常，可能的原因：\n" +
                    "1. Microsoft Visio未正确安装或注册\n" +
                    "2. Visio进程被锁定或权限不足\n" +
                    "3. 系统COM组件损坏\n" +
                    $"详细错误: {ex.Message}");
            }
            catch (Exception ex)
            {
                ReportLog($"转换出错: {ex.Message}");
                ReportLog($"异常类型: {ex.GetType().Name}");
                if (ex.InnerException != null)
                {
                    ReportLog($"内部异常: {ex.InnerException.Message}");
                }
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

        /// <summary>
        /// 检查Visio是否可用
        /// </summary>
        /// <returns>检查结果</returns>
        public ConversionResult CheckVisioAvailability()
        {
            Microsoft.Office.Interop.Visio.Application? visioApp = null;
            try
            {
                ReportLog("正在检查Visio环境...");
                
                // 尝试创建Visio应用程序
                visioApp = new Microsoft.Office.Interop.Visio.Application();
                var version = visioApp.Version;
                ReportLog($"✅ Visio可用，版本: {version}");
                return ConversionResult.Success(new string[] { $"Visio版本: {version}" });
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ReportLog($"❌ Visio不可用: {ex.Message}");
                return ConversionResult.Error($"Visio环境检查失败：\n" +
                    "1. 请确认Microsoft Visio已正确安装\n" +
                    "2. 检查Visio是否已正确注册\n" +
                    "3. 尝试手动启动Visio测试\n" +
                    $"错误详情: {ex.Message}");
            }
            catch (Exception ex)
            {
                ReportLog($"❌ 环境检查异常: {ex.Message}");
                return ConversionResult.Error($"环境检查失败: {ex.Message}");
            }
            finally
            {
                // 清理COM对象
                if (visioApp != null)
                {
                    try
                    {
                        visioApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(visioApp);
                    }
                    catch
                    {
                        // 忽略清理时的异常
                    }
                }
            }
        }

        private void ReportProgress(int percentage, string message)
        {
            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs(percentage, message));
        }

        private void ReportLog(string message)
        {
            LogMessage?.Invoke(this, new ConversionLogEventArgs(DateTime.Now, message));
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // 释放托管资源
            }

            // 释放Visio COM对象
            if (_appConfig != null)
            {
                _appConfig.Dispose();
                _appConfig = null;
                // 重置静态实例
                AppConfig.Instance = new AppConfig();
            }
            _disposed = true;
        }

        ~ConversionService()
        {
            Dispose(false);
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