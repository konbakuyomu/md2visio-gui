using md2visio.main;
using md2visio.struc.figure;

namespace md2visio.main
{
    /// <summary>
    /// 控制台应用程序入口点
    /// </summary>
    public static class ConsoleApp
    {
        public static void Main(string[] args)
        {
            try
            {
                var config = new AppConfig();
                if (config.ParseArgs(args))
                {
                    config.Main();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"错误: {ex.Message}");
                Environment.Exit(1);
            }
        }
    }
} 