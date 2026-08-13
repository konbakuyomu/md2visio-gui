using md2visio.GUI.Services;
using md2visio.GUI.Localization;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;

namespace md2visio.GUI.Forms
{
    /// <summary>
    /// md2visio 主窗口
    /// </summary>
    public partial class MainForm : Form
    {
        private readonly ConversionService _conversionService;

        // 控件声明
        private Panel _dragDropPanel = null!;
        private Label _dragDropLabel = null!;
        private Label _selectedFileLabel = null!;
        private TextBox _outputDirTextBox = null!;
        private TextBox _fileNameTextBox = null!;
        private CheckBox _showVisioCheckBox = null!;
        private CheckBox _silentOverwriteCheckBox = null!;
        private RichTextBox _logTextBox = null!;
        private ProgressBar _progressBar = null!;
        private Label _statusLabel = null!;
        private Button _browseFileButton = null!;
        private Button _selectDirButton = null!;
        private Button _startConversionButton = null!;
        private Button _openOutputButton = null!;
        private Button _clearLogButton = null!;
        private ComboBox _languageComboBox = null!;
        private bool _changingLanguage;


        private string? _selectedFilePath;

        public MainForm()
        {
            _conversionService = new ConversionService();
            _conversionService.ProgressChanged += OnProgressChanged;
            _conversionService.LogMessage += OnLogMessage;

            InitializeComponent();
            SetupEventHandlers();
            UpdateUI();
        }

        private void InitializeComponent()
        {
            // 窗口设置
            Text = UiStrings.Get("AppTitle");
            Size = new Size(1250, 850);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(600, 500);

            // 创建主面板
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new Padding(10)
            };

            // 设置行高比例
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 标题
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 120)); // 文件选择区域
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 120)); // 输出设置
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80)); // 选项
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 支持类型
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // 日志区域
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80)); // 按钮和状态栏

            Controls.Add(mainPanel);

            // 创建各个区域
            CreateTitleArea(mainPanel, 0);
            CreateFileSelectionArea(mainPanel, 1);
            CreateOutputSettingsArea(mainPanel, 2);
            CreateOptionsArea(mainPanel, 3);
            CreateSupportedTypesArea(mainPanel, 4);
            CreateLogArea(mainPanel, 5);
            CreateStatusArea(mainPanel, 6);
        }

        private void CreateTitleArea(TableLayoutPanel parent, int row)
        {
            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

            var titleLabel = new Label
            {
                Text = UiStrings.Get("HeaderTitle"),
                Font = UiFont(12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var languageLabel = new Label
            {
                Text = UiStrings.Get("Language"),
                AutoSize = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(8, 0, 5, 0)
            };

            _languageComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _languageComboBox.Items.Add(new LanguageChoice("en", UiStrings.Get("English")));
            _languageComboBox.Items.Add(new LanguageChoice("zh-CN", UiStrings.Get("Chinese")));
            _languageComboBox.SelectedIndex = CultureSettings.CurrentCultureName == "zh-CN" ? 1 : 0;
            _languageComboBox.SelectedIndexChanged += OnLanguageChanged;

            container.Controls.Add(titleLabel, 0, 0);
            container.Controls.Add(languageLabel, 1, 0);
            container.Controls.Add(_languageComboBox, 2, 0);
            parent.Controls.Add(container, 0, row);
        }

        private void CreateFileSelectionArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = UiStrings.Get("InputFile"),
                Dock = DockStyle.Fill,
                Font = UiFont(9, FontStyle.Bold)
            };

            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(10)
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 30));

            // 拖拽区域
            _dragDropPanel = new Panel
            {
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.LightGray,
                Dock = DockStyle.Fill,
                AllowDrop = true
            };

            _dragDropLabel = new Label
            {
                Text = UiStrings.Get("DropHint"),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = UiFont(10)
            };
            _dragDropPanel.Controls.Add(_dragDropLabel);

            // 浏览按钮
            _browseFileButton = new Button
            {
                Text = UiStrings.Get("BrowseFiles"),
                Dock = DockStyle.Fill,
                Font = UiFont(9),
                Margin = new Padding(10, 0, 0, 0)
            };

            // 选中文件显示
            _selectedFileLabel = new Label
            {
                Text = UiStrings.Get("NoFileSelected"),
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                Font = UiFont(8)
            };

            container.Controls.Add(_dragDropPanel, 0, 0);
            container.Controls.Add(_browseFileButton, 1, 0);
            container.Controls.Add(_selectedFileLabel, 0, 1);
            container.SetColumnSpan(_selectedFileLabel, 2);

            groupBox.Controls.Add(container);
            parent.Controls.Add(groupBox, 0, row);
        }

        private void CreateOutputSettingsArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = UiStrings.Get("OutputSettings"),
                Dock = DockStyle.Fill,
                Font = UiFont(9, FontStyle.Bold)
            };

            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2,
                Padding = new Padding(10, 10, 10, 10)
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));

            // 输出目录
            var outputDirLabel = new Label { Text = UiStrings.Get("OutputDirectory"), TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill, Font = SystemFonts.MessageBoxFont };
            _outputDirTextBox = new TextBox { Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), Dock = DockStyle.Fill, Font = UiFont(9) };
            _selectDirButton = new Button { Text = UiStrings.Get("SelectDirectory"), Dock = DockStyle.Fill, Margin = new Padding(5, 0, 0, 0), Font = SystemFonts.MessageBoxFont };

            // 文件名
            var fileNameLabel = new Label { Text = UiStrings.Get("FileName"), TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill, Font = SystemFonts.MessageBoxFont };
            _fileNameTextBox = new TextBox { Text = "output", Dock = DockStyle.Fill, Font = UiFont(9) };
            var extensionLabel = new Label { Text = ".vsdx", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, Font = UiFont(9) };

            container.Controls.Add(outputDirLabel, 0, 0);
            container.Controls.Add(_outputDirTextBox, 1, 0);
            container.Controls.Add(_selectDirButton, 2, 0);
            container.Controls.Add(fileNameLabel, 0, 1);
            container.Controls.Add(_fileNameTextBox, 1, 1);
            container.Controls.Add(extensionLabel, 2, 1);

            groupBox.Controls.Add(container);
            parent.Controls.Add(groupBox, 0, row);
        }

        private void CreateOptionsArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = UiStrings.Get("ConversionOptions"),
                Dock = DockStyle.Fill,
                Font = UiFont(9, FontStyle.Bold)
            };

            var container = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10, 20, 10, 20),
                WrapContents = false
            };

            _showVisioCheckBox = new CheckBox
            {
                Text = UiStrings.Get("ShowVisio"),
                AutoSize = true,
                Font = UiFont(9),
                Margin = new Padding(0, 0, 30, 0)
            };

            _silentOverwriteCheckBox = new CheckBox
            {
                Text = UiStrings.Get("SilentOverwrite"),
                AutoSize = true,
                Font = UiFont(9),
                Checked = true
            };

            container.Controls.Add(_showVisioCheckBox);
            container.Controls.Add(_silentOverwriteCheckBox);

            groupBox.Controls.Add(container);
            parent.Controls.Add(groupBox, 0, row);
        }

        private void CreateSupportedTypesArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = UiStrings.Get("SupportedTypes"),
                Dock = DockStyle.Top,
                Font = UiFont(9, FontStyle.Bold),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            var container = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(10, 15, 10, 15),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };

            // 创建单个类型标签
            var supportedTypes = new[]
            {
                (UiStrings.Get("Flowchart"), "graph/flowchart"),
                (UiStrings.Get("PieChart"), "pie"),
                (UiStrings.Get("UserJourney"), "journey"),
                (UiStrings.Get("PacketDiagram"), "packet"),
                (UiStrings.Get("XYChart"), "xychart"),
                (UiStrings.Get("SequenceDiagram"), "sequence"),
                (UiStrings.Get("EntityRelationshipDiagram"), "er")
            };

            foreach (var (icon, name) in supportedTypes)
            {
                var label = new Label
                {
                    Text = $"{icon} {name}",
                    AutoSize = true,
                    Font = UiFont(9),
                    ForeColor = icon.StartsWith("✅") ? Color.DarkGreen : Color.Red,
                    Margin = new Padding(0, 5, 15, 5)
                };
                container.Controls.Add(label);
            }

            groupBox.Controls.Add(container);
            parent.Controls.Add(groupBox, 0, row);

            void SyncSupportedTypesWidth()
            {
                // FlowLayoutPanel 需要受限宽度才能正确计算换行后的高度
                int width = groupBox.ClientSize.Width - container.Margin.Horizontal - container.Padding.Horizontal;
                if (width > 0)
                    container.MaximumSize = new Size(width, 0);
            }

            groupBox.SizeChanged += (_, __) => SyncSupportedTypesWidth();
            groupBox.HandleCreated += (_, __) => SyncSupportedTypesWidth();
        }

        private void CreateLogArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = UiStrings.Get("ConversionLog"),
                Dock = DockStyle.Fill,
                Font = UiFont(9, FontStyle.Bold)
            };

            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(5)
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));

            _logTextBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.Black,
                ForeColor = Color.Lime
            };

            _clearLogButton = new Button
            {
                Text = UiStrings.Get("ClearLog"),
                Dock = DockStyle.Fill,
                Font = UiFont(9),
                Margin = new Padding(5, 5, 0, 5),
                MinimumSize = new Size(85, 30)
            };

            container.Controls.Add(_logTextBox, 0, 0);
            container.Controls.Add(_clearLogButton, 1, 0);

            groupBox.Controls.Add(container);
            parent.Controls.Add(groupBox, 0, row);
        }

        private void CreateStatusArea(TableLayoutPanel parent, int row)
        {
            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 2
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 按钮
            _startConversionButton = new Button
            {
                Text = UiStrings.Get("StartConversion"),
                Dock = DockStyle.Fill,
                BackColor = Color.LightGreen,
                Font = UiFont(9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };

            var checkVisioButton = new Button
            {
                Text = UiStrings.Get("CheckVisio"),
                Dock = DockStyle.Fill,
                BackColor = Color.LightBlue,
                Font = UiFont(9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };
            checkVisioButton.Click += OnCheckVisioClick;

            _openOutputButton = new Button
            {
                Text = UiStrings.Get("OpenOutput"),
                Dock = DockStyle.Fill,
                Enabled = false,
                Margin = new Padding(0, 0, 5, 0)
            };

            var exitButton = new Button
            {
                Text = UiStrings.Get("Exit"),
                Dock = DockStyle.Fill,
                BackColor = Color.LightCoral,
                Margin = new Padding(0, 0, 5, 0)
            };
            exitButton.Click += (s, e) => Close();

            // 状态标签
            _statusLabel = new Label
            {
                Text = UiStrings.Get("Ready"),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = UiFont(9)
            };

            // 进度条
            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Visible = false
            };

            container.Controls.Add(_startConversionButton, 0, 0);
            container.Controls.Add(checkVisioButton, 1, 0);
            container.Controls.Add(_openOutputButton, 2, 0);
            container.Controls.Add(exitButton, 3, 0);
            container.Controls.Add(_statusLabel, 4, 0);

            var authorLabel = new LinkLabel
            {
                Text = "© konbakuyomu",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight,
                Font = UiFont(9)
            };
            authorLabel.Links.Add(0, authorLabel.Text.Length, "https://github.com/konbakuyomu/md2visio-gui/");
            authorLabel.LinkClicked += (s, e) => {
                if (e.Link?.LinkData is string url)
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            };

            container.Controls.Add(authorLabel, 5, 0);
            container.Controls.Add(_progressBar, 0, 1);
            container.SetColumnSpan(_progressBar, 6);

            parent.Controls.Add(container, 0, row);
        }

        private void SetupEventHandlers()
        {
            // 拖拽事件
            _dragDropPanel.DragEnter += OnDragEnter;
            _dragDropPanel.DragDrop += OnDragDrop;
            _dragDropPanel.Click += OnDragPanelClick;

            // 按钮事件
            _browseFileButton.Click += OnBrowseFileClick;
            _selectDirButton.Click += OnSelectDirClick;
            _startConversionButton.Click += OnStartConversionClick;
            _openOutputButton.Click += OnOpenOutputClick;
            _clearLogButton.Click += OnClearLogClick;

            // 文件名自动更新
            _selectedFileLabel.TextChanged += OnSelectedFileChanged;
        }

        private void OnDragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            {
                e.Effect = DragDropEffects.Copy;
                _dragDropPanel.BackColor = Color.LightBlue;
            }
        }

        private void OnDragDrop(object? sender, DragEventArgs e)
        {
            _dragDropPanel.BackColor = Color.LightGray;
            
            if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            {
                var file = files[0];
                if (Path.GetExtension(file).Equals(".md", StringComparison.OrdinalIgnoreCase))
                {
                    SetSelectedFile(file);
                }
                else
                {
                    MessageBox.Show(UiStrings.Get("InvalidMarkdown"), UiStrings.Get("Error"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void OnDragPanelClick(object? sender, EventArgs e)
        {
            OnBrowseFileClick(sender, e);
        }

        private void OnBrowseFileClick(object? sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog
            {
                Filter = UiStrings.Get("MarkdownFilter"),
                Title = UiStrings.Get("SelectMarkdown")
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                SetSelectedFile(dialog.FileName);
            }
        }

        private void OnSelectDirClick(object? sender, EventArgs e)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = UiStrings.Get("SelectOutputFolder"),
                SelectedPath = _outputDirTextBox.Text
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _outputDirTextBox.Text = dialog.SelectedPath;
            }
        }

        private async void OnStartConversionClick(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedFilePath))
            {
                MessageBox.Show(UiStrings.Get("ChooseInputFirst"), UiStrings.Get("Notice"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrEmpty(_outputDirTextBox.Text))
            {
                MessageBox.Show(UiStrings.Get("ChooseOutputFirst"), UiStrings.Get("Notice"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SetUIBusy(true);

            try
            {
                var result = await _conversionService.ConvertAsync(
                    _selectedFilePath,
                    _outputDirTextBox.Text,
                    _fileNameTextBox.Text, // 传递用户设置的文件名
                    _showVisioCheckBox.Checked,
                    _silentOverwriteCheckBox.Checked
                );

                if (result.IsSuccess)
                {
                    _openOutputButton.Enabled = true;
                    ShowUserMessage(
                        UiStrings.Format("ConversionSucceeded", result.OutputFiles?.Length ?? 0),
                        UiStrings.Get("Success"),
                        MessageBoxIcon.Information);
                }
                else
                {
                    ShowUserMessage(
                        UiStrings.Format("ConversionFailed", result.ErrorMessage),
                        UiStrings.Get("Error"),
                        MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                ShowUserMessage(
                    UiStrings.Format("ConversionException", ex.Message),
                    UiStrings.Get("Error"),
                    MessageBoxIcon.Error);
            }
            finally
            {
                SetUIBusy(false);
            }
        }

        private void OnOpenOutputClick(object? sender, EventArgs e)
        {
            if (Directory.Exists(_outputDirTextBox.Text))
            {
                Process.Start("explorer.exe", _outputDirTextBox.Text);
            }
        }

        private void OnClearLogClick(object? sender, EventArgs e)
        {
            _logTextBox.Clear();
        }

        private void OnSelectedFileChanged(object? sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedFilePath))
            {
                var fileName = Path.GetFileNameWithoutExtension(_selectedFilePath);
                _fileNameTextBox.Text = fileName;
            }
        }

        private void SetSelectedFile(string filePath)
        {
            _selectedFilePath = filePath;
            _selectedFileLabel.Text = UiStrings.Format("SelectedFile", filePath);
            _selectedFileLabel.ForeColor = Color.Green;

            // 检测图表类型
            var types = _conversionService.DetectMermaidTypes(filePath);
            if (types.Count > 0)
            {
                LogMessage(UiStrings.Format("DetectedTypes", string.Join(", ", types)));
            }

            UpdateUI();
        }

        private void SetUIBusy(bool busy)
        {
            _startConversionButton.Enabled = !busy;
            _browseFileButton.Enabled = !busy;
            _selectDirButton.Enabled = !busy;
            _progressBar.Visible = busy;
            
            if (busy)
            {
                _statusLabel.Text = UiStrings.Get("Converting");
                _progressBar.Value = 0;
            }
            else
            {
                _statusLabel.Text = UiStrings.Get("Ready");
            }
        }

        private void UpdateUI()
        {
            _startConversionButton.Enabled = !string.IsNullOrEmpty(_selectedFilePath);
        }

        private void OnProgressChanged(object? sender, ConversionProgressEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => OnProgressChanged(sender, e)));
                return;
            }

            _progressBar.Value = e.Percentage;
            _statusLabel.Text = e.Message;
        }

        private void OnLogMessage(object? sender, ConversionLogEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => OnLogMessage(sender, e)));
                return;
            }

            LogMessage($"[{e.Timestamp:HH:mm:ss}] {e.Message}");
        }

        private void LogMessage(string message)
        {
            _logTextBox.AppendText($"{message}\n");
            _logTextBox.ScrollToCaret();
        }

        private void ShowUserMessage(string message, string caption, MessageBoxIcon icon)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }
            Activate();
            BringToFront();
            MessageBox.Show(this, message, caption, MessageBoxButtons.OK, icon);
        }

        private async void OnCheckVisioClick(object? sender, EventArgs e)
        {
            SetUIBusy(true);
            _statusLabel.Text = UiStrings.Get("CheckingVisio");

            try
            {
                var result = await Task.Run(() => _conversionService.CheckVisioAvailability());
                
                if (result.IsSuccess)
                {
                    MessageBox.Show(UiStrings.Format("VisioCheckPassed", string.Join("\n", result.OutputFiles ?? [])),
                        UiStrings.Get("VisioCheckSucceededTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _statusLabel.Text = UiStrings.Get("VisioAvailable");
                }
                else
                {
                    MessageBox.Show(UiStrings.Format("VisioCheckFailed", result.ErrorMessage),
                        UiStrings.Get("VisioCheckFailedTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _statusLabel.Text = UiStrings.Get("VisioUnavailable");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(UiStrings.Format("CheckException", ex.Message), UiStrings.Get("ExceptionTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                _statusLabel.Text = UiStrings.Get("CheckFailed");
            }
            finally
            {
                SetUIBusy(false);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 释放服务持有的资源，例如Visio COM对象
            _conversionService.Dispose();
            base.OnFormClosing(e);
        }

        private void OnLanguageChanged(object? sender, EventArgs e)
        {
            if (_changingLanguage || _languageComboBox.SelectedItem is not LanguageChoice choice ||
                choice.CultureName == CultureSettings.CurrentCultureName)
                return;

            var outputDirectory = _outputDirTextBox.Text;
            var fileName = _fileNameTextBox.Text;
            var showVisio = _showVisioCheckBox.Checked;
            var silentOverwrite = _silentOverwriteCheckBox.Checked;
            var log = _logTextBox.Text;

            _changingLanguage = true;
            CultureSettings.SaveAndApply(choice.CultureName);
            SuspendLayout();
            var oldControls = Controls.Cast<Control>().ToArray();
            Controls.Clear();
            foreach (var control in oldControls)
                control.Dispose();
            InitializeComponent();
            SetupEventHandlers();
            _outputDirTextBox.Text = outputDirectory;
            _fileNameTextBox.Text = fileName;
            _showVisioCheckBox.Checked = showVisio;
            _silentOverwriteCheckBox.Checked = silentOverwrite;
            _logTextBox.Text = log;
            if (!string.IsNullOrEmpty(_selectedFilePath))
            {
                _selectedFileLabel.Text = UiStrings.Format("SelectedFile", _selectedFilePath);
                _selectedFileLabel.ForeColor = Color.Green;
            }
            UpdateUI();
            ResumeLayout(true);
            _changingLanguage = false;
        }

        private sealed record LanguageChoice(string CultureName, string DisplayName)
        {
            public override string ToString() => DisplayName;
        }

        private static Font UiFont(float size, FontStyle style = FontStyle.Regular) =>
            new((SystemFonts.MessageBoxFont ?? SystemFonts.DefaultFont).FontFamily, size, style);
    }
} 
