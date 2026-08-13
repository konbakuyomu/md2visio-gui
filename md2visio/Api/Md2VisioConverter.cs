using md2visio.mermaid.cmn;
using md2visio.Localization;
using md2visio.struc.figure;
using md2visio.vsdx.@base;
using System.Reflection;

namespace md2visio.Api
{
    /// <summary>
    /// Mermaid 到 Visio 转换器实现
    /// 包装现有转换逻辑，提供简洁的 API
    /// </summary>
    public sealed class Md2VisioConverter : IMd2VisioConverter
    {
        private IVisioSession? _session;
        private bool _disposed;
        private readonly object _lock = new object();

        /// <summary>
        /// 执行转换
        /// </summary>
        public ConversionResult Convert(
            ConversionRequest request,
            IProgress<ConversionProgress>? progress = null,
            ILogSink? logger = null)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(Md2VisioConverter));

            logger ??= NullLogSink.Instance;

            try
            {
                return ConvertInternal(request, progress, logger);
            }
            catch (NotImplementedException ex)
            {
                logger.Error(CoreStrings.Format("UnsupportedDiagram", ex.Message));
                return ConversionResult.Failed(CoreStrings.Format("UnsupportedDiagram", ex.Message), ex);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                logger.Error(CoreStrings.Format("VisioComErrorDetail", ex.Message));
                return ConversionResult.Failed(
                    CoreStrings.Get("VisioComError"),
                    ex);
            }
            catch (Exception ex)
            {
                var root = UnwrapException(ex);
                if (!ReferenceEquals(root, ex))
                {
                    logger.Error(CoreStrings.Format("ConversionFailedTyped", root.GetType().Name, root.Message));
                    return ConversionResult.Failed(CoreStrings.Format("ConversionFailedTyped", root.GetType().Name, root.Message), ex);
                }

                logger.Error(CoreStrings.Format("ConversionFailed", ex.Message));
                return ConversionResult.Failed(CoreStrings.Format("ConversionFailed", ex.Message), ex);
            }
        }

        private static Exception UnwrapException(Exception ex)
        {
            Exception current = ex;
            if (current is TargetInvocationException tie && tie.InnerException != null)
            {
                current = tie.InnerException;
            }

            while (current.InnerException != null)
            {
                current = current.InnerException;
            }

            return current;
        }

        private ConversionResult ConvertInternal(
            ConversionRequest request,
            IProgress<ConversionProgress>? progress,
            ILogSink logger)
        {
            // Step 1: 验证输入
            progress?.Report(new ConversionProgress(0, CoreStrings.Get("ValidatingInput"), ConversionPhase.Starting));
            logger.Info(CoreStrings.Format("InputFile", request.InputPath));
            logger.Info(CoreStrings.Format("OutputPath", request.OutputPath));

            if (!File.Exists(request.InputPath))
            {
                return ConversionResult.Failed(CoreStrings.Format("InputMissing", request.InputPath));
            }

            if (!Path.GetExtension(request.InputPath).Equals(".md", StringComparison.OrdinalIgnoreCase))
            {
                return ConversionResult.Failed(CoreStrings.Get("InputMustBeMarkdown"));
            }

            // Fail before Visio startup/rendering when an explicitly named output file
            // cannot be overwritten. The final save repeats this check to cover races.
            if (request.SilentOverwrite &&
                request.OutputPath.EndsWith(".vsdx", StringComparison.OrdinalIgnoreCase))
            {
                var outputStatus = OutputFileAccess.Check(request.OutputPath);
                if (outputStatus != OutputFileStatus.Writable)
                    return ConversionResult.Failed(OutputFileAccess.GetMessage(outputStatus, request.OutputPath));
            }

            // 确保输出目录存在
            string? outputDir = request.OutputPath.EndsWith(".vsdx", StringComparison.OrdinalIgnoreCase)
                ? Path.GetDirectoryName(request.OutputPath)
                : request.OutputPath;

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
                logger.Debug(CoreStrings.Format("CreatedOutputDirectory", outputDir));
            }

            // Step 2: 创建转换上下文和 Visio 会话
            progress?.Report(new ConversionProgress(20, CoreStrings.Get("InitializingVisio"), ConversionPhase.Starting));
            logger.Info(CoreStrings.Get("InitializingContext"));

            var context = new ConversionContext(request, logger);
            _session = new VisioSession(request.ShowVisio);

            // Step 3: 解析 Mermaid 内容
            progress?.Report(new ConversionProgress(30, CoreStrings.Get("ParsingMermaid"), ConversionPhase.Parsing));
            logger.Info(CoreStrings.Get("ParsingMermaid"));

            var synContext = new SynContext(request.InputPath);
            SttMermaidStart.Run(synContext);

            if (request.Debug)
            {
                logger.Debug(synContext.ToString());
            }

            // Step 4: 构建图表
            progress?.Report(new ConversionProgress(50, CoreStrings.Get("BuildingDiagram"), ConversionPhase.Building));
            logger.Info(CoreStrings.Get("BuildingDiagram"));

            var factory = new FigureBuilderFactory(synContext.NewSttIterator(), context, _session);

            // Step 5: 渲染到 Visio
            progress?.Report(new ConversionProgress(70, CoreStrings.Get("RenderingVisio"), ConversionPhase.Rendering));
            logger.Info(CoreStrings.Get("RenderingVisioFormat"));

            factory.Build(request.OutputPath);

            if (!string.IsNullOrWhiteSpace(context.LastError))
            {
                if (!request.ShowVisio)
                {
                    _session?.Dispose();
                    _session = null;
                }
                return ConversionResult.Failed(context.LastError);
            }

            // Step 6: 如果不显示 Visio 则清理
            progress?.Report(new ConversionProgress(90, CoreStrings.Get("SavingOutput"), ConversionPhase.Saving));

            if (!request.ShowVisio)
            {
                _session.Dispose();
                _session = null;
            }

            // Step 7: 收集输出文件并提供详细反馈
            progress?.Report(new ConversionProgress(100, CoreStrings.Get("ConversionComplete"), ConversionPhase.Completed));

            var outputFiles = CollectOutputFiles(request);

            if (outputFiles.Length > 0)
            {
                logger.Info(CoreStrings.Format("GeneratedFiles", outputFiles.Length));
                foreach (var file in outputFiles)
                {
                    logger.Info($"  - {Path.GetFileName(file)}");
                }
                return ConversionResult.Succeeded(outputFiles);
            }
            else
            {
                // 提供详细的错误原因
                if (factory.FiguresBuilt == 0)
                {
                    var supportedTypes = string.Join(", ", TypeMap.BuilderMap.Keys.Distinct().OrderBy(k => k));
                    if (factory.UnsupportedTypes.Count > 0)
                    {
                        var unsupported = string.Join(", ", factory.UnsupportedTypes);
                        return ConversionResult.Failed(CoreStrings.Format("UnsupportedTypesInFile", unsupported, supportedTypes));
                    }
                    else
                    {
                        return ConversionResult.Failed(CoreStrings.Get("NoValidDiagram"));
                    }
                }

                logger.Warning(CoreStrings.Get("NoOutputWarning"));
                return ConversionResult.Failed(CoreStrings.Get("NoOutputError"));
            }
        }

        /// <summary>
        /// 收集输出文件
        /// </summary>
        private string[] CollectOutputFiles(ConversionRequest request)
        {
            if (request.OutputPath.EndsWith(".vsdx", StringComparison.OrdinalIgnoreCase))
            {
                // 文件模式：检查指定文件
                return File.Exists(request.OutputPath)
                    ? new[] { request.OutputPath }
                    : Array.Empty<string>();
            }
            else
            {
                // 目录模式：查找所有 .vsdx 文件
                return Directory.Exists(request.OutputPath)
                    ? Directory.GetFiles(request.OutputPath, "*.vsdx")
                    : Array.Empty<string>();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;

            lock (_lock)
            {
                if (_disposed) return;

                _session?.Dispose();
                _session = null;

                _disposed = true;
            }
        }
    }
}
