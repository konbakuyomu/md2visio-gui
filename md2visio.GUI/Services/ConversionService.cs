using md2visio.Api;
using md2visio.Localization;

namespace md2visio.GUI.Services
{
    /// <summary>
    /// Mermaid 到 Visio 转换服务 - 使用新的 API 层
    /// </summary>
    public class ConversionService : IDisposable
    {
        public event EventHandler<ConversionProgressEventArgs>? ProgressChanged;
        public event EventHandler<ConversionLogEventArgs>? LogMessage;

        private IMd2VisioConverter? _converter;
        private bool _disposed = false;
        private bool _automationTimedOut;
        private readonly object _lock = new object();
        private readonly object _logFileLock = new object();
        private static readonly TimeSpan VisioStartupTimeout = TimeSpan.FromSeconds(30);

        public string LogFilePath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "md2visio",
            "logs",
            $"md2visio-{DateTime.Now:yyyyMMdd}.log");

        /// <summary>
        /// 转换 MD 文件到 Visio（异步）
        /// </summary>
        public async Task<ConversionResult> ConvertAsync(
            string inputFile,
            string outputDir,
            string? fileName = null,
            bool showVisio = false,
            bool silentOverwrite = false)
        {
            if (_automationTimedOut)
                return ConversionResult.Error(CoreStrings.Get("RestartAfterTimeout"));

            var startupCompleted = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var conversionTask = RunOnStaThread(() => Convert(
                inputFile,
                outputDir,
                fileName,
                showVisio,
                silentOverwrite,
                startupCompleted));

            var startupOrCompletion = Task.WhenAny(conversionTask, startupCompleted.Task);
            if (await Task.WhenAny(startupOrCompletion, Task.Delay(VisioStartupTimeout)) != startupOrCompletion)
            {
                _automationTimedOut = true;
                var message = CoreStrings.Format("VisioStartupTimeout", (int)VisioStartupTimeout.TotalSeconds);
                ReportLog(message);
                return ConversionResult.Error(message);
            }

            return await conversionTask;
        }

        private Task<ConversionResult> RunOnStaThread(Func<ConversionResult> work)
        {
            var tcs = new TaskCompletionSource<ConversionResult>(TaskCreationOptions.RunContinuationsAsynchronously);

            var thread = new Thread(() =>
            {
                try
                {
                    tcs.SetResult(work());
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            })
            {
                IsBackground = true
            };

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        /// <summary>
        /// 同步转换方法
        /// </summary>
        private ConversionResult Convert(
            string inputFile,
            string outputDir,
            string? fileName,
            bool showVisio,
            bool silentOverwrite,
            TaskCompletionSource<bool> startupCompleted)
        {
            try
            {
                ReportProgress(0, CoreStrings.Get("StartingConversion"));
                ReportLog(CoreStrings.Format("InputFile", inputFile));
                ReportLog(CoreStrings.Format("OutputDirectory", outputDir));

                // 验证输入文件
                if (!File.Exists(inputFile))
                    return ConversionResult.Error(CoreStrings.Format("InputMissing", inputFile));

                if (!Path.GetExtension(inputFile).Equals(".md", StringComparison.OrdinalIgnoreCase))
                    return ConversionResult.Error(CoreStrings.Get("InputMustBeMarkdown"));

                // 创建输出目录
                Directory.CreateDirectory(outputDir);
                ReportProgress(10, CoreStrings.Get("PreparingEnvironment"));

                // 构建输出路径
                string outputPath = BuildOutputPath(outputDir, fileName);
                ReportLog(CoreStrings.Format("OutputPath", outputPath));

                // 创建转换请求
                var request = new md2visio.Api.ConversionRequest(
                    inputPath: inputFile,
                    outputPath: outputPath,
                    showVisio: showVisio,
                    silentOverwrite: silentOverwrite,
                    debug: false
                );

                // 创建进度报告器
                var progress = new InlineProgress<md2visio.Api.ConversionProgress>(p =>
                {
                    int guiProgress = MapProgress(p.Phase, p.Percentage);
                    ReportProgress(guiProgress, p.Message);
                    if (p.Phase >= md2visio.Api.ConversionPhase.Parsing)
                        startupCompleted.TrySetResult(true);
                });

                // 创建日志接收器
                var logSink = new GuiLogSink(this);

                // 获取或创建转换器
                lock (_lock)
                {
                    if (showVisio)
                    {
                        // 显示模式：复用转换器以保持 Visio 窗口
                        _converter ??= new md2visio.Api.Md2VisioConverter();
                    }
                    else
                    {
                        // 非显示模式：每次创建新的转换器
                        _converter?.Dispose();
                        _converter = new md2visio.Api.Md2VisioConverter();
                    }
                }

                // 执行转换
                ReportLog(CoreStrings.Get("ExecutingCore"));
                var apiResult = _converter.Convert(request, progress, logSink);
                if (apiResult.Success)
                    ReportLog(CoreStrings.Get("CoreComplete"));

                // 非显示模式立即释放资源
                if (!showVisio)
                {
                    lock (_lock)
                    {
                        _converter?.Dispose();
                        _converter = null;
                    }
                }

                // 转换结果
                if (apiResult.Success)
                {
                    ReportProgress(100, CoreStrings.Get("ConversionComplete"));
                    ReportLog(CoreStrings.Format("GeneratedFiles", apiResult.OutputFiles.Length));
                    foreach (var file in apiResult.OutputFiles)
                    {
                        ReportLog($"  - {Path.GetFileName(file)}");
                    }
                    return ConversionResult.Success(apiResult.OutputFiles);
                }
                else
                {
                    ReportLog(CoreStrings.Format("ConversionFailed", apiResult.ErrorMessage));
                    return ConversionResult.Error(apiResult.ErrorMessage ?? CoreStrings.Get("UnknownError"));
                }
            }
            catch (NotImplementedException ex)
            {
                ReportLog(CoreStrings.Format("FeatureNotImplemented", ex.Message));
                return ConversionResult.Error(CoreStrings.Format("UnsupportedDiagram", ex.Message));
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ReportLog(CoreStrings.Format("ComException", ex.Message, ex.HResult.ToString("X8")));
                return ConversionResult.Error(CoreStrings.Format("ComExceptionHelp", ex.Message));
            }
            catch (Exception ex)
            {
                ReportLog(CoreStrings.Format("ConversionFailed", ex.Message));
                ReportLog(CoreStrings.Format("ErrorType", ex.GetType().Name));
                if (ex.InnerException != null)
                {
                    ReportLog(CoreStrings.Format("InnerException", ex.InnerException.Message));
                }
                return ConversionResult.Error(CoreStrings.Format("ConversionFailed", ex.Message));
            }
        }

        /// <summary>
        /// 构建输出路径
        /// </summary>
        private string BuildOutputPath(string outputDir, string? fileName)
        {
            if (!string.IsNullOrEmpty(fileName))
            {
                if (!fileName.EndsWith(".vsdx", StringComparison.OrdinalIgnoreCase))
                    fileName += ".vsdx";
                return Path.Combine(outputDir, fileName);
            }
            return outputDir;
        }

        /// <summary>
        /// 映射 API 进度到 GUI 进度
        /// </summary>
        private int MapProgress(md2visio.Api.ConversionPhase phase, int apiPercentage)
        {
            return phase switch
            {
                md2visio.Api.ConversionPhase.Starting => 10,
                md2visio.Api.ConversionPhase.Parsing => 30,
                md2visio.Api.ConversionPhase.Building => 50,
                md2visio.Api.ConversionPhase.Rendering => 70,
                md2visio.Api.ConversionPhase.Saving => 90,
                md2visio.Api.ConversionPhase.Completed => 100,
                _ => apiPercentage
            };
        }

        /// <summary>
        /// 检测 MD 文件中的 Mermaid 图表类型
        /// </summary>
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
                ReportLog(CoreStrings.Format("DetectTypesError", ex.Message));
            }

            return types;
        }

        /// <summary>
        /// 检查 Visio 是否可用
        /// </summary>
        public ConversionResult CheckVisioAvailability()
        {
            Microsoft.Office.Interop.Visio.Application? visioApp = null;
            try
            {
                ReportLog(CoreStrings.Get("CheckingVisio"));

                visioApp = new Microsoft.Office.Interop.Visio.Application();
                var version = visioApp.Version;
                ReportLog(CoreStrings.Format("VisioAvailableVersion", version));
                return ConversionResult.Success([CoreStrings.Format("VisioVersion", version)]);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                ReportLog(CoreStrings.Format("VisioUnavailableDetail", ex.Message));
                return ConversionResult.Error(CoreStrings.Format("VisioCheckHelp", ex.Message));
            }
            catch (Exception ex)
            {
                ReportLog(CoreStrings.Format("EnvironmentCheckException", ex.Message));
                return ConversionResult.Error(CoreStrings.Format("EnvironmentCheckFailed", ex.Message));
            }
            finally
            {
                if (visioApp != null)
                {
                    try
                    {
                        visioApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(visioApp);
                    }
                    catch { }
                }
            }
        }

        private void ReportProgress(int percentage, string message)
        {
            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs(percentage, message));
        }

        private void ReportLog(string message)
        {
            var timestamp = DateTime.Now;
            WriteLogFile(timestamp, message);
            LogMessage?.Invoke(this, new ConversionLogEventArgs(timestamp, message));
        }

        private void WriteLogFile(DateTime timestamp, string message)
        {
            try
            {
                lock (_logFileLock)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath)!);
                    File.AppendAllText(LogFilePath, $"[{timestamp:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}");
                }
            }
            catch
            {
                // File logging must never interrupt conversion or UI logging.
            }
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
                lock (_lock)
                {
                    _converter?.Dispose();
                    _converter = null;
                }
            }

            _disposed = true;
        }

        ~ConversionService()
        {
            Dispose(false);
        }

        /// <summary>
        /// GUI 日志接收器 - 将 API 日志转发到 GUI 事件
        /// </summary>
        private class GuiLogSink : md2visio.Api.ILogSink
        {
            private readonly ConversionService _service;
            private static readonly string[] LevelPrefixes = new[] { "[DEBUG]", "[WARN]", "[ERROR]", "[INFO]" };

            public GuiLogSink(ConversionService service)
            {
                _service = service;
            }

            public void Info(string message) => _service.ReportLog(message);
            public void Debug(string message) => _service.ReportLog(WithPrefix("[DEBUG]", message));
            public void Warning(string message) => _service.ReportLog(WithPrefix("[WARN]", message));
            public void Error(string message) => _service.ReportLog(WithPrefix("[ERROR]", message));

            private static string WithPrefix(string prefix, string message)
            {
                if (string.IsNullOrWhiteSpace(message))
                {
                    return prefix;
                }

                foreach (var levelPrefix in LevelPrefixes)
                {
                    if (message.StartsWith(levelPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        return message;
                    }
                }

                return $"{prefix} {message}";
            }
        }

        private sealed class InlineProgress<T>(Action<T> handler) : IProgress<T>
        {
            public void Report(T value) => handler(value);
        }
    }

    /// <summary>
    /// 转换结果（保持向后兼容）
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
