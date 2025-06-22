using md2visio.main;
using Microsoft.Win32;
using Visio = Microsoft.Office.Interop.Visio;

namespace md2visio.vsdx.@base
{
    internal abstract class VBuilder
    {
        public static Visio.Application? VisioApp = null;
        protected Visio.Document visioDoc;
        protected Visio.Page visioPage;
        
        public VBuilder()
        {
            EnsureVisioApp();
            visioDoc = VisioApp!.Documents.Add(""); // 添加一个新文档
            visioPage = visioDoc.Pages[1]; // 获取活动页面
        }

        /// <summary>
        /// 确保Visio应用程序可用，如果不可用则重新创建
        /// </summary>
        private static void EnsureVisioApp()
        {
            try
            {
                // 检查现有VisioApp是否可用
                if (VisioApp != null)
                {
                    // 尝试访问Version属性来测试COM对象是否有效
                    _ = VisioApp.Version;
                    return; // 如果没有异常，说明COM对象仍然有效
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // COM对象已经无效，需要重新创建
                Console.WriteLine($"COM异常，重新创建Visio应用程序: {ex.Message}");
                VisioApp = null;
            }
            catch (System.Runtime.InteropServices.InvalidComObjectException ex)
            {
                // COM对象已被释放，需要重新创建
                Console.WriteLine($"COM对象已释放，重新创建: {ex.Message}");
                VisioApp = null;
            }

            try
            {
                // 创建新的Visio应用程序
                Console.WriteLine("正在创建Visio应用程序...");
                VisioApp = new Visio.Application();
                VisioApp.Visible = AppConfig.Instance.Visible;
                Console.WriteLine($"Visio应用程序创建成功，版本: {VisioApp.Version}");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                throw new ApplicationException($"无法创建Visio应用程序。请确认：\n" +
                    $"1. Microsoft Visio已正确安装\n" +
                    $"2. 当前用户有权限访问Visio\n" +
                    $"3. Visio未被其他进程锁定\n" +
                    $"错误详情: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"创建Visio应用程序时发生未知错误: {ex.Message}", ex);
            }
        }

        public void SaveAndClose(string outputFile)
        {
            visioPage.ResizeToFitContents();

            AppConfig config = AppConfig.Instance;
            
            // 如果显示Visio窗口，给用户时间查看结果
            if (config.Visible && VisioApp != null)
            {
                VisioApp.Visible = true; // 确保可见
                // 可以添加延迟让用户查看绘制结果
                System.Threading.Thread.Sleep(1000); // 1秒延迟
            }

            bool overwrite = true;
            if (!config.Quiet && File.Exists(outputFile))
            {
                Console.WriteLine($"Output file '{outputFile}' exists, input Y to overwrite: ");
                overwrite = Console.ReadLine()?.ToLower() == "y";
            }

            if (overwrite) visioDoc.SaveAsEx(outputFile, 0);
            else visioDoc.Saved = true;

            // 如果不显示Visio，立即关闭；如果显示，保持打开状态
            if (!config.Visible)
            {
            visioDoc.Close();
            }
            else
            {
                // 显示模式下，让用户手动关闭或保持打开状态
                visioDoc.Saved = true; // 标记为已保存，避免提示
            }
        }


        public static string? GetVisioContentDirectory()
        {
            // 支持从Office 2003 (11.0) 到 Office 2019/2021 (16.0)
            int[] officeVersions = Enumerable.Range(11, 16).ToArray();

            foreach (int version in officeVersions)
            {
                string subKey = $@"Software\Microsoft\Office\{version}.0\Visio\InstallRoot";
#pragma warning disable CA1416, CS8604
                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(subKey);
                object? value = key?.GetValue("Path");
                if (value != null)
                {
                    string contentDir = Path.Combine(value.ToString(), "Visio Content");
#pragma warning restore CA1416, CS8604
                    if (Directory.Exists(contentDir))
                    {
                        return contentDir;
                    }
                }
            }

            return null;
        }
    }
}
