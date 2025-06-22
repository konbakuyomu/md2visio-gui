using md2visio.GUI.Services;
using System.ComponentModel;
using System.Diagnostics;

namespace md2visio.GUI.Forms
{
    /// <summary>
    /// md2visio 主窗口
    /// </summary>
    public partial class MainForm : Form
    {
        private readonly ConversionService _conversionService;
        private BackgroundWorker? _backgroundWorker;

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
        private Label _supportedTypesLabel = null!;

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
            Text = "md2visio - Mermaid 转 Visio 工具";
            Size = new Size(800, 700);
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
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 80)); // 输出设置
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); // 选项
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); // 支持类型
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
            var titleLabel = new Label
            {
                Text = "📄 md2visio - Mermaid 转 Visio 工具",
                Font = new Font("Microsoft YaHei UI", 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            parent.Controls.Add(titleLabel, 0, row);
        }

        private void CreateFileSelectionArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = "📁 输入文件",
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
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
                Text = "将 .md 文件拖拽到此处或点击浏览选择",
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 10)
            };
            _dragDropPanel.Controls.Add(_dragDropLabel);

            // 浏览按钮
            _browseFileButton = new Button
            {
                Text = "浏览文件...",
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9),
                Margin = new Padding(10, 0, 0, 0)
            };

            // 选中文件显示
            _selectedFileLabel = new Label
            {
                Text = "未选择文件",
                Dock = DockStyle.Fill,
                ForeColor = Color.Gray,
                Font = new Font("Microsoft YaHei UI", 8)
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
                Text = "📂 输出设置",
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
            };

            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 2,
                Padding = new Padding(10)
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            // 输出目录
            var outputDirLabel = new Label { Text = "输出目录:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill };
            _outputDirTextBox = new TextBox { Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), Dock = DockStyle.Fill };
            _selectDirButton = new Button { Text = "选择目录...", Dock = DockStyle.Fill, Margin = new Padding(5, 0, 0, 0) };

            // 文件名
            var fileNameLabel = new Label { Text = "文件名:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill };
            _fileNameTextBox = new TextBox { Text = "output", Dock = DockStyle.Fill };
            var extensionLabel = new Label { Text = ".vsdx", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };

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
                Text = "⚙️ 转换选项",
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
            };

            var container = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(10)
            };

            _showVisioCheckBox = new CheckBox
            {
                Text = "显示 Visio 窗口",
                AutoSize = true,
                Margin = new Padding(0, 0, 20, 0)
            };

            _silentOverwriteCheckBox = new CheckBox
            {
                Text = "静默覆盖文件",
                AutoSize = true,
                Checked = true
            };

            container.Controls.Add(_showVisioCheckBox);
            container.Controls.Add(_silentOverwriteCheckBox);

            groupBox.Controls.Add(container);
            parent.Controls.Add(groupBox, 0, row);
        }

        private void CreateSupportedTypesArea(TableLayoutPanel parent, int row)
        {
            _supportedTypesLabel = new Label
            {
                Text = "📊 支持的图表类型: ✅ 流程图(graph/flowchart)  ✅ 饼图(pie)  ✅ 用户旅程图(journey)  ✅ 数据包图(packet)  ✅ XY图表(xychart)  ❌ 时序图(未实现)",
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 8),
                ForeColor = Color.DarkGreen,
                TextAlign = ContentAlignment.MiddleLeft
            };

            parent.Controls.Add(_supportedTypesLabel, 0, row);
        }

        private void CreateLogArea(TableLayoutPanel parent, int row)
        {
            var groupBox = new GroupBox
            {
                Text = "📝 转换日志",
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold)
            };

            var container = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(5)
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));

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
                Text = "清空日志",
                Dock = DockStyle.Top,
                Height = 30,
                Margin = new Padding(5, 0, 0, 0)
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
                ColumnCount = 4,
                RowCount = 2
            };
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            container.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            container.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            container.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // 按钮
            _startConversionButton = new Button
            {
                Text = "🚀 开始转换",
                Dock = DockStyle.Fill,
                BackColor = Color.LightGreen,
                Font = new Font("Microsoft YaHei UI", 9, FontStyle.Bold),
                Margin = new Padding(0, 0, 5, 0)
            };

            _openOutputButton = new Button
            {
                Text = "📁 打开输出目录",
                Dock = DockStyle.Fill,
                Enabled = false,
                Margin = new Padding(0, 0, 5, 0)
            };

            var exitButton = new Button
            {
                Text = "❌ 退出",
                Dock = DockStyle.Fill,
                BackColor = Color.LightCoral,
                Margin = new Padding(0, 0, 5, 0)
            };
            exitButton.Click += (s, e) => Close();

            // 状态标签
            _statusLabel = new Label
            {
                Text = "就绪",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9)
            };

            // 进度条
            _progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Visible = false
            };

            container.Controls.Add(_startConversionButton, 0, 0);
            container.Controls.Add(_openOutputButton, 1, 0);
            container.Controls.Add(exitButton, 2, 0);
            container.Controls.Add(_statusLabel, 3, 0);
            container.Controls.Add(_progressBar, 0, 1);
            container.SetColumnSpan(_progressBar, 4);

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
                    MessageBox.Show("请选择 .md 文件！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                Filter = "Markdown 文件|*.md|所有文件|*.*",
                Title = "选择 Markdown 文件"
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
                Description = "选择输出目录",
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
                MessageBox.Show("请先选择要转换的文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrEmpty(_outputDirTextBox.Text))
            {
                MessageBox.Show("请选择输出目录！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SetUIBusy(true);

            try
            {
                var result = await _conversionService.ConvertAsync(
                    _selectedFilePath,
                    _outputDirTextBox.Text,
                    _showVisioCheckBox.Checked,
                    _silentOverwriteCheckBox.Checked
                );

                if (result.IsSuccess)
                {
                    _openOutputButton.Enabled = true;
                    MessageBox.Show($"转换成功！\n生成了 {result.OutputFiles?.Length} 个文件。", "成功", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"转换失败！\n错误: {result.ErrorMessage}", "错误", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"转换过程中发生错误:\n{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            _selectedFileLabel.Text = $"选中文件: {filePath}";
            _selectedFileLabel.ForeColor = Color.Green;

            // 检测图表类型
            var types = _conversionService.DetectMermaidTypes(filePath);
            if (types.Count > 0)
            {
                LogMessage($"检测到图表类型: {string.Join(", ", types)}");
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
                _statusLabel.Text = "转换中...";
                _progressBar.Value = 0;
            }
            else
            {
                _statusLabel.Text = "就绪";
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

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_backgroundWorker?.IsBusy == true)
            {
                var result = MessageBox.Show("转换正在进行中，确定要退出吗？", "确认", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
            }

            base.OnFormClosing(e);
        }
    }
} 