## md2visio

> 一个 .NET C# 应用，旨在将 Markdown 文件中嵌入的 Mermaid.js 语法转换为 Microsoft Visio 图形。

> 本项目的目的是实现从 Mermaid 文本描述到 Visio 可视化图表的自动化转换，从而提高技术文档和系统设计中图表制作的效率与一致性。

> **已完成**: 同时提供命令行版本和GUI图形界面版本，满足不同用户需求
> **架构特点**: 采用类库+GUI的架构设计，核心逻辑作为类库被GUI项目引用，实现代码复用

> (暂无团队信息)

> .NET 8.0, C# (核心类库 + Windows Forms GUI版本)

## Dependencies

*   **YamlDotNet**: 用于解析 `default` 目录下的 YAML 配置文件，加载图表样式和布局。
*   **WonderCircuits.Microsoft.Office.Interop.Visio**: 核心依赖，用于与本地安装的 Microsoft Visio 应用程序进行 COM 交互，实现图形的程序化绘制。
    - **COM线程模式**: GUI版本确保STA线程模式，兼容Windows Forms和Visio COM交互
    - **错误处理**: 实现了详细的COM异常诊断和用户友好的错误提示
    - **环境检测**: 支持预先检测Visio是否可用，避免运行时错误
*   **System.Drawing.Common**: 用于图形计算和文本度量。
*   **Windows Forms**: GUI版本基于Windows Forms框架，提供现代化用户界面。

## Development Environment

*   **IDE**: Visual Studio 2022 (或兼容的 .NET IDE)
*   **SDK**: .NET 8.0 SDK
*   **软件**: Microsoft Visio (需要本地安装)
*   **运行方式**: GUI版本可直接运行，控制台版本通过特定配置编译为可执行文件
*   **发布模式**: GUI版本支持单文件自包含发布，生成独立的可执行文件

## 部署和分发

*   **发布包结构**: 包含 md2visio.GUI.exe、配置文件夹、模板文件和用户文档
*   **模板文件**: md2visio.vssx 采用智能路径查找，自动适配开发和发布环境
*   **用户体验**: 提供安装助手 (install.bat) 和详细的中文使用说明
*   **系统要求**: Windows 10/11 x64 + Microsoft Visio，无需安装 .NET 运行时

## Architecture

> 项目采用分层架构设计，核心功能作为类库，支持多种前端调用方式。

### 项目类型配置
*   **md2visio**: 默认编译为类库 (Library)，供其他项目引用
*   **md2visio.GUI**: 编译为Windows Forms可执行文件 (WinExe)
*   **特殊配置**: md2visio在'ConsoleRelease'配置下可编译为控制台应用

### 核心架构分层
1. **解析层** (mermaid): Mermaid语法解析器
2. **数据层** (struc): 图形数据结构 (AST)
3. **绘制层** (vsdx): Visio绘制引擎
4. **界面层** (GUI): Windows Forms用户界面
5. **服务层** (Services): 转换服务和业务逻辑

## Structrue

> 项目结构清晰，遵循功能模块化的设计原则。核心逻辑分为解析(mermaid)、数据结构(struc)和绘制(vsdx)三大部分，实现了关注点分离。

## GUI特性增强 (v2025.06.22)

*   **Visio环境检查**: 新增"🔍 检查Visio"功能，支持预先验证Visio环境
*   **增强错误处理**: 提供详细的COM异常诊断和解决建议
*   **线程模式优化**: 确保正确的COM线程模式，提高稳定性
*   **用户体验**: 异步操作、实时反馈、友好的错误提示和问题排查指引
*   **COM对象生命周期**: 优化了Visio COM对象的生命周期管理，通过IDisposable模式确保在“显示Visio窗口”时进程稳定，解决了闪退问题。

```
.
├── md2visio.sln          # Visual Studio 解决方案文件
├── md2visio.vssx         # Visio 模板文件
├── README.md             # 项目说明文档
├── GUI使用指南.md        # GUI版本使用指南
├── md2visio/             # **核心类库项目** 🔄
│   ├── md2visio.csproj # C# 项目文件，配置为类库模式
│   ├── default/        # 存放各类图表的默认 YAML 配置文件
│   │   ├── *.yaml      # 不同图表类型 (flowchart, pie, etc.) 的样式配置
│   │   └── theme/      # 主题相关的样式配置
│   ├── main/           # 主程序入口与配置
│   │   ├── AppConfig.cs  # **应用配置类** (public，支持GUI调用)
│   │   ├── AppTest.cs    # 测试入口
│   │   └── ConsoleApp.cs # **控制台应用入口** (用于ConsoleRelease配置)
└── md2visio.GUI/         # **GUI项目** 🆕
    ├── md2visio.GUI.csproj # GUI项目配置文件，引用md2visio类库
    ├── Forms/            # 窗体文件
    │   └── MainForm.cs   # **主窗口** - 完整的GUI界面实现
    ├── Services/         # 服务层
    │   └── ConversionService.cs # **转换服务** - 封装转换逻辑，提供事件通知
    ├── Resources/        # 资源文件夹
    └── Program.cs        # **GUI入口点** - 启动Windows Forms应用
    ├── mermaid/        # **Mermaid 解析器模块**
    │   ├── @cmn/       # 解析器通用组件，实现状态机核心框架
    │   │   ├── SynContext.cs # 解析器上下文，管理输入流和状态
    │   │   ├── SynState.cs   # 状态机抽象基类
    │   │   └── ...       # 其他具体的状态实现和工具
    │   ├── graph/      # 图表 (Graph) 解析逻辑
    │   ├── journey/    # 用户旅程图 (Journey) 解析逻辑
    │   ├── packet/     # 包图 (Packet) 解析逻辑
    │   ├── pie/        # 饼图 (Pie) 解析逻辑
    │   └── xy/         # XY图 (XY Chart) 解析逻辑
    ├── struc/          # **图形数据结构 (AST) 模块**
    │   ├── figure/     # 通用图形元素基类 (Figure, Node, Edge)
    │   │   ├── Figure.cs # 所有图表数据结构的基类
    │   │   └── FigureBuilder.cs # 构建数据结构对象的抽象工厂
    │   ├── graph/      # Graph 的数据结构 (Graph, GNode, GEdge)
    │   ├── journey/    # Journey 的数据结构
    │   ├── packet/     # Packet 的数据结构
    │   ├── pie/        # Pie 的数据结构
    │   └── xy/         # XY Chart 的数据结构
    ├── test/           # 存放用于测试的 .md 文件
    └── vsdx/           # **Visio 绘制模块**
        ├── @base/      # 绘制器的基类和通用接口
        │   ├── VBuilder.cs # Visio 构建器基类，负责计算布局
        │   ├── VDrawer.cs  # Visio 绘制器基类，负责执行绘图API调用
        │   └── ...
        ├── @tool/      # 颜色等绘图辅助工具
        ├── VBuilderG.cs  # Graph 的 Visio 构建器
        ├── VDrawerG.cs   # Graph 的 Visio 绘制器
        ├── VBuilderJo.cs # Journey 的 Visio 构建器
        ├── VDrawerJo.cs  # Journey 的 Visio 绘制器
        └── ... (其他图表类型的 Builder 和 Drawer)
