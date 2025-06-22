using md2visio.GUI.Forms;

namespace md2visio.GUI;

internal static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // 确保COM线程模式
        System.Threading.Thread.CurrentThread.SetApartmentState(System.Threading.ApartmentState.STA);
        
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        
        // 设置应用程序外观
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        
        // 启动主窗口
        Application.Run(new MainForm());
    }    
}