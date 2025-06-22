## md2visio - Mermaid 到 Visio 的现代化转换工具

> 本项目是一个基于 .NET 8 和 C# 12 开发的桌面应用程序，旨在提供一个强大、稳定且用户友好的 Mermaid.js 到 Microsoft Visio 的转换解决方案。它不仅包含了原版 `md2visio` 的所有核心转换逻辑，还通过一个全新的图形用户界面 (GUI) 对其进行了彻底的现代化改造，极大地提升了易用性和稳定性。

> **核心价值**: 解决大语言模型 (LLM) 生成的 Mermaid 图表难以在专业绘图工具中进行二次编辑的痛点，打通从 AI 生成到专业精修的工作流。

> **项目状态**: **功能完善，体验优化**。同时提供稳定的核心转换库 (`md2visio`) 和一个功能完备的图形界面 (`md2visio.GUI`)。

> .NET 8.0, C# 12, Windows Forms

## Dependencies

*   **Microsoft.Office.Interop.Visio**: **核心依赖**。通过COM互操作性，实现与本地安装的Microsoft Visio应用程序的深度集成和程序化控制。
*   **YamlDotNet**: 用于解析 `default` 目录下的 YAML 配置文件，为不同类型的图表提供灵活的样式和主题定制能力。
*   **System.Drawing.Common**: 用于进行图形计算和文本尺寸度量，是精确布局的基础。
*   **Windows Forms**: GUI版本基于成熟的 Windows Forms 框架构建，确保在 Windows 系统下的原生体验和良好兼容性。

## Architecture

> 项目采用现代化的分层架构，实现了核心逻辑与用户界面的完全分离，具有良好的可维护性和扩展性。

### 核心架构分层
1.  **展现层 (GUI)**: 基于 `Windows Forms` 构建，通过 `md2visio.GUI/Forms/MainForm.cs` 提供所有用户交互功能。
2.  **服务层 (Services)**: `md2visio.GUI/Services/ConversionService.cs` 作为UI与核心逻辑的桥梁，封装了转换流程、异步处理和 COM 对象生命周期管理。
3.  **应用层 (main)**: `md2visio/main/AppConfig.cs` 负责处理命令行参数和全局配置，同时服务于控制台和GUI两种模式。
4.  **解析层 (mermaid)**: 位于 `md2visio/mermaid/`，采用**状态机模式**对 Mermaid 源码进行词法和语法分析。
5.  **数据结构层 (struc)**: 位于 `md2visio/struc/`，定义了图表的抽象语法树 (AST)，作为解析层和绘制层之间的数据契约。
6.  **绘制层 (vsdx)**: 位于 `md2visio/vsdx/`，通过 COM 互操作调用 Visio API，将 AST 精确地绘制到 Visio 画布上。

## 详细项目结构 (Source Code)

```
.
├── md2visio.sln              # Visual Studio 解决方案文件
├── md2visio.vssx             # Visio 模板文件，包含预定义的形状和样式
├── README.md                 # 项目介绍与使用指南
│
├── md2visio/                 # ================== 核心逻辑库 ==================
│   ├── md2visio.csproj       # C# 项目文件，配置为类库 (Library)
│   │
│   ├── main/                 # --- 应用主干与配置 ---
│   │   ├── AppConfig.cs      # 全局配置类，管理命令行参数和应用状态
│   │   ├── AppTest.cs        # 用于内部测试的入口类
│   │   └── ConsoleApp.cs     # 控制台版本的程序主入口
│   │
│   ├── mermaid/              # --- Mermaid 解析器模块 (状态机) ---
│   │   ├── @cmn/             # 通用解析器组件
│   │   │   ├── SynContext.cs # 解析器上下文，管理文本流和状态
│   │   │   ├── SynState.cs   # 所有状态类的抽象基类
│   │   │   └── ... (其他状态机基础组件)
│   │   ├── graph/            # 流程图 (graph) 的解析逻辑和状态实现
│   │   ├── journey/          # 用户旅程图 (journey) 的解析逻辑
│   │   ├── packet/           # 数据包图 (packet) 的解析逻辑
│   │   ├── pie/              # 饼图 (pie) 的解析逻辑
│   │   └── xy/               # XY图 (xy) 的解析逻辑
│   │
│   ├── struc/                # --- 图形数据结构 (AST) ---
│   │   ├── figure/           # 通用图形元素定义
│   │   │   ├── Figure.cs     # 所有图表数据结构的基类
│   │   │   ├── FigureBuilder.cs # 构建AST的抽象工厂基类
│   │   │   ├── Node.cs       # 节点的抽象表示
│   │   │   ├── Edge.cs       # 连接线的抽象表示
│   │   │   └── Config.cs     # 图表样式的配置模型
│   │   ├── graph/            # 流程图的数据结构 (Graph, GNode, GEdge)
│   │   ├── journey/          # 用户旅程图的数据结构 (Journey, JoSection, JoTask)
│   │   ├── packet/           # 数据包图的数据结构 (Packet, PacBlock)
│   │   ├── pie/              # 饼图的数据结构 (Pie, PieDataItem)
│   │   └── xy/               # XY图的数据结构 (XyChart, XyAxis)
│   │
│   └── vsdx/                 # --- Visio 绘制引擎 ---
│       ├── @base/            # 绘制器的基类和通用接口
│       │   ├── VBuilder.cs   # Visio 构建器基类，负责计算布局
│       │   ├── VDrawer.cs    # Visio 绘制器基类，负责执行绘图API调用
│       │   ├── VShapeDrawer.cs # 封装了对Visio形状的底层操作
│       │   └── ... (其他绘制器基础组件)
│       ├── @tool/            # 颜色、坐标等绘图辅助工具
│       │   └── VColor.cs     # 颜色处理工具类
│       ├── VBuilderG.cs      # 流程图的 Visio 布局构建器
│       ├── VDrawerG.cs       # 流程图的 Visio 绘制器
│       ├── VBuilderJo.cs     # 用户旅程图的 Visio 布局构建器
│       ├── VDrawerJo.cs      # 用户旅程图的 Visio 绘制器
│       ├── VBuilderPac.cs    # 数据包图的 Visio 布局构建器
│       ├── VDrawerPac.cs     # 数据包图的 Visio 绘制器
│       ├── VBuilderPie.cs    # 饼图的 Visio 布局构建器
│       ├── VDrawerPie.cs     # 饼图的 Visio 绘制器
│       ├── VBuilderXy.cs     # XY图的 Visio 布局构建器
│       └── VDrawerXy.cs      # XY图的 Visio 绘制器
│
└── md2visio.GUI/             # ================== 图形用户界面 ==================
    ├── md2visio.GUI.csproj   # GUI项目配置文件，引用 md2visio 核心库
    ├── Program.cs            # GUI 应用程序的入口点，负责启动主窗口
    ├── Forms/                # --- 窗体文件 ---
    │   └── MainForm.cs       # 主窗口，包含所有UI控件的布局和事件处理逻辑
    └── Services/             # --- 服务层 ---
        └── ConversionService.cs # 转换服务，封装核心库调用、异步处理和COM对象生命周期管理
