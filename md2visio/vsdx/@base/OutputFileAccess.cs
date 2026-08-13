using md2visio.Localization;

namespace md2visio.vsdx.@base;

internal enum OutputFileStatus
{
    Writable,
    InUse,
    AccessDenied
}

internal static class OutputFileAccess
{
    public static OutputFileStatus Check(string outputFile)
    {
        if (!File.Exists(outputFile)) return OutputFileStatus.Writable;

        try
        {
            using var stream = new FileStream(
                outputFile,
                FileMode.Open,
                FileAccess.ReadWrite,
                FileShare.None);
            return OutputFileStatus.Writable;
        }
        catch (UnauthorizedAccessException)
        {
            return OutputFileStatus.AccessDenied;
        }
        catch (IOException)
        {
            return OutputFileStatus.InUse;
        }
    }

    public static string GetMessage(OutputFileStatus status, string outputFile) => status switch
    {
        OutputFileStatus.InUse => CoreStrings.Format("OutputFileInUse", outputFile),
        OutputFileStatus.AccessDenied => CoreStrings.Format("OutputFileAccessDenied", outputFile),
        _ => string.Empty
    };
}
