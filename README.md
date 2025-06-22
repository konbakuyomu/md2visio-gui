# md2visio-gui - 一款更易用的 Mermaid 转 Visio 工具

![GUI界面截图](https://raw.githubusercontent.com/konbakuyomu/md2visio-gui/main/assets/gui-screenshot.png)  <!-- 请将此处的URL替换为您自己的截图URL -->

这是一个基于 .NET 8 和 Windows Forms 的桌面应用程序，它可以将您用 Mermaid.js 语法编写的图表，轻松转换为 Microsoft Visio 的 `.vsdx` 格式文件。

与原版相比，本项目最大的特点是提供了一个直观的图形用户界面 (GUI)，让不熟悉命令行的朋友也能愉快地使用。

## ✨ 项目缘起与致谢

本项目是在 [Megre/md2visio](https://github.com/Megre/md2visio) 这个优秀项目的基础上进行二次开发的。

原项目为 Mermaid 到 Visio 的转换提供了强大的核心逻辑，解决了最关键的技术难题。我在它的基础上，主要做了以下工作：
*   **开发了全新的图形用户界面 (GUI)**，让操作更直观、更简单。
*   **修复了若干稳定性问题**，例如在特定情况下 Visio 进程闪退的 bug。
*   **优化了UI布局和用户体验**，让软件用起来更顺手。
*   **重构了部分代码**，使其更易于维护和扩展。

在此，特别感谢原作者 **Megre** 的杰出工作和开源贡献！

## 🚀 主要功能

*   **图形化操作**: 告别命令行，所有功能都可以在窗口里点几下鼠标完成。
*   **拖拽支持**: 直接把 `.md` 文件拖进程序窗口，自动加载。
*   **实时日志**: 黑底绿字的日志窗口，实时显示转换的每一步，方便排查问题。
*   **灵活的输出设置**: 可以自由指定输出的文件夹和文件名。
*   **Visio 显示控制**: 你可以选择在转换时，实时看着 Visio 窗口画图；也可以让它在后台默默完成。
*   **环境自检**: 不确定自己的电脑环境行不行？点一下“检查Visio”按钮，程序会帮你判断。

## 📊 支持的 Mermaid 图表类型

这是当前版本对 Mermaid 图表的支持情况。我们会持续努力支持更多类型！

- [x] **graph / flowchart** (流程图)
  - [x] themes (支持主题)
- [x] **journey** (用户旅程图)
  - [x] themes (支持主题)
- [x] **pie** (饼图)
  - [x] themes (支持主题)
- [x] **packet-beta** (数据包图)
  - [x] themes (支持主题)
- [x] **xychart-beta** (XY图表)
- [x] **Configuration** (配置指令)
  - [x] `frontmatter`
  - [x] `directive`
- [ ] **sequenceDiagram** (时序图) - **暂不支持**
- [ ] **classDiagram** (类图) - **暂不支持**
- [ ] **stateDiagram / stateDiagram-v2** (状态图) - **暂不支持**
- [ ] **erDiagram** (实体关系图) - **暂不支持**
- [ ] **gantt** (甘特图) - **暂不支持**
- [ ] **quadrantChart** (象限图) - **暂不支持**
- [ ] **requirementDiagram** (需求图) - **暂不支持**
- [ ] **gitGraph** (Git图) - **暂不支持**
- [ ] **C4Context** (C4上下文图) - **暂不支持**
- [ ] **mindmap** (脑图) - **暂不支持**
- [ ] **timeline** (时间轴) - **暂不支持**
- [ ] **zenuml** - **暂不支持**
- [ ] **sankey-beta** (桑基图) - **暂不支持**
- [ ] **block-beta** (块状图) - **暂不支持**
- [ ] **kanban** (看板) - **暂不支持**
- [ ] **architecture-beta** (架构图) - **暂不支持**


## 🛠️ 使用指南 (给普通用户)

1.  **下载**: 前往本项目的 [Releases](https://github.com/konbakuyomu/md2visio-gui/releases) 页面，下载最新版本的 `md2visio-gui-win-x64.zip` 文件。
2.  **解压**: 将下载的压缩包解压到你电脑上的任意位置。
3.  **‼️ 重要前提 ‼️**: 确保你的电脑上已经安装了 **Microsoft Visio** 桌面版。这是程序运行的必要条件。
4.  **运行**: 双击解压出来的 `md2visio.GUI.exe` 文件，即可启动程序。

## 👨‍💻 二次开发指南 (给开发者)

如果你想对这个项目进行修改或贡献代码，请遵循以下步骤：

### **环境要求**
*   **Visual Studio 2022**: 建议使用最新版本。
*   **.NET 8.0 SDK**: 确保已安装。
*   **Microsoft Visio**: 开发和调试时需要。

### **项目结构**
本项目主要由两个工程组成：
*   `md2visio/`: **核心逻辑库**。这是一个 `.NET` 类库项目，包含了所有 Mermaid 语法的解析、数据结构转换和 Visio 绘图的核心代码。
*   `md2visio.GUI/`: **图形用户界面**。这是一个 `Windows Forms` 项目，它引用了 `md2visio` 核心库，并为其提供了一个用户友好的图形界面。

### **如何编译**
1.  使用 Visual Studio 2022 打开 `md2visio.sln` 解决方案文件。
2.  将解决方案配置设置为 `Debug` 或 `Release`。
3.  直接生成解决方案即可。

### **核心代码导览**
*   **Mermaid 解析器**: 位于 `md2visio/mermaid/` 目录下，采用状态机模式，对不同类型的图表进行逐行解析。
*   **图形数据结构**: 位于 `md2visio/struc/` 目录下，这是解析器和绘图器之间的桥梁，定义了图表的抽象语法树 (AST)。
*   **Visio 绘制引擎**: 位于 `md2visio/vsdx/` 目录下，通过 `Microsoft.Office.Interop.Visio` COM 组件与 Visio 应用程序交互，负责在画布上创建形状、连接线和设置样式。
*   **GUI 服务层**: `md2visio.GUI/Services/ConversionService.cs` 封装了对核心库的调用，并处理了 COM 对象的生命周期管理，是 GUI 与后端逻辑交互的枢纽。

欢迎提交 Pull Request 或在 Issues 中提出你的想法！