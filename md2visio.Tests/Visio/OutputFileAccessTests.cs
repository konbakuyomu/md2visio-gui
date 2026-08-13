using md2visio.vsdx.@base;

namespace md2visio.Tests.VisioParsing;

public sealed class OutputFileAccessTests
{
    [Fact]
    public void Check_ReturnsWritableForMissingFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"md2visio-{Guid.NewGuid():N}.vsdx");

        Assert.Equal(OutputFileStatus.Writable, OutputFileAccess.Check(path));
    }

    [Fact]
    public void Check_ReturnsInUseForExclusivelyLockedFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"md2visio-{Guid.NewGuid():N}.vsdx");
        File.WriteAllText(path, "test");

        try
        {
            using (var lockStream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                Assert.Equal(OutputFileStatus.InUse, OutputFileAccess.Check(path));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }
}
