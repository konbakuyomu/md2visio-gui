namespace md2visio.main
{
    public class AppConfig : IDisposable
    {
        public static AppConfig Instance { get; set; } = new AppConfig();

        public string InputFile { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public bool Visible { get; set; } = false;
        public bool Quiet { get; set; } = false;
        public bool Debug { get; set; } = false;
        public bool Test { get; set; } = false;

        private md2visio.struc.figure.FigureBuilderFactory? _factory;
        private bool _disposed = false;

        public AppConfig() { }

        public bool ParseArgs(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i].ToUpper();
                if(arg=="/I" && i + 1 < args.Length)
                    InputFile = args[++i];
                else if(arg=="/O" && i + 1 < args.Length)
                    OutputPath = args[++i];
                else if(arg=="/Y") 
                    Quiet = true;
                else if(arg=="/V")
                    Visible = true;
                else if (arg == "/D")
                    Debug = true;
                else if (arg == "/T")
                    Test = true;
                else if(arg == "/?" 
                    || arg == "/H"   || arg == "/HELP"
                    || arg == "-H"   || arg == "-HELP"
                    || arg == "--H"  || arg == "--HELP")
                    ShowHelp();
            }

            bool success = InputFile != string.Empty && 
                OutputPath != string.Empty;
            if(!success) ShowHelp();

            return success;
        }

        public bool LoadArguments(string[] args)
        {
            return ParseArgs(args);
        }

        public void Main()
        {
            // 执行转换逻辑（从原Program.cs复制）
            try
            {
                var synContext = new md2visio.mermaid.cmn.SynContext(InputFile);
                md2visio.mermaid.cmn.SttMermaidStart.Run(synContext);

                if (Debug)
                    Console.Write(synContext.ToString());

                _factory ??= new md2visio.struc.figure.FigureBuilderFactory(synContext.NewSttIterator());
                _factory.Build(OutputPath);

                // 如果不是Visio可见模式，则立即退出
                if (!Visible)
                {
                    _factory.Quit();
                    _factory = null;
                }
            }
            catch (Exception ex)
            {
                if (Debug)
                    throw;
                else
                    throw new ApplicationException(ex.Message, ex);
            }
        }

        void ShowHelp()
        {
            Console.WriteLine(
                "md2visio /I MD_FILE /O OUTPUT_PATH [/V] [/Y]\n\n" +
                "/I\tSpecify the path of the input file (.md)\n" +
                "/O\tSpecify the path/folder of the output file (.vsdx)\n" +
                "/V\tShow the Visio application while drawing (optional, default is invisible)\n" +
                "/Y\tQuietly overwrite the existing output file (optional, by default requires user confirmation" +
                "/?\tPrint this help");
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

            // 释放COM对象
            _factory?.Quit();
            _factory = null;

            _disposed = true;
        }

        ~AppConfig()
        {
            Dispose(false);
        }
    }
}
