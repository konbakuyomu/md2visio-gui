# Changelog

## 🐛 GUI闪退问题修复 (2025-06-22) - 关键稳定性修复

### 🎯 问题定位
- **核心原因**: 在GUI中启用“显示Visio窗口”选项时，程序会立即闪退。经分析，根本原因在于`Visio.Application` COM对象的生命周期管理不当。该对象在`ConversionService`中被创建为局部变量，导致方法执行完毕后被GC过早回收，从而使Visio进程终止。

### 🔧 修复方案
- **生命周期管理**: 采用`IDisposable`模式对COM对象的生命周期进行显式管理。
  - **`ConversionService`**: 实现`IDisposable`接口，并持有一个`AppConfig`实例，确保在需要时复用同一个Visio进程。
  - **`AppConfig`**: 实现`IDisposable`接口，负责管理其内部的`FigureBuilderFactory`（进而管理Visio实例）的创建与销毁。
  - **`MainForm`**: 在窗体关闭事件 (`OnFormClosing`) 中调用`_conversionService.Dispose()`，确保应用程序退出时能安全地释放所有COM资源，关闭Visio进程。

### 📁 文件变更记录
- 修改: `md2visio.GUI/Services/ConversionService.cs` (实现IDisposable，持有AppConfig实例)
- 修改: `md2visio/main/AppConfig.cs` (实现IDisposable，管理Factory生命周期)
- 修改: `md2visio.GUI/Forms/MainForm.cs` (在窗体关闭时调用Dispose)

### ✅ 修复效果
- **问题解决**: 彻底解决了启用“显示Visio窗口”时的闪退问题。
- **稳定性提升**: 确保了Visio COM对象在整个GUI应用的生命周期内被正确管理，提高了程序的健壮性。

---
##  GUI Visio兼容性修复 (2025-06-22) - 关键问题解决

### 🎯 问题分析与解决
- **根本问题识别**: 分析发现GUI版本无法正常打开Visio的核心原因：
  1. COM线程模型不正确 - Windows Forms需要STA线程模式
  2. 错误处理不足 - 缺乏对COM异常的详细诊断
  3. 环境检测缺失 - 没有预先验证Visio是否可用

### 🔧 技术修复实施
- **COM线程模式优化**: 
  - 在GUI程序入口添加 `[STAThread]` 属性
  - 显式设置 `Thread.CurrentThread.SetApartmentState(ApartmentState.STA)`
- **增强错误处理**: 
  - 在 `VBuilder.cs` 中增加详细的Visio应用程序创建诊断信息
  - 在 `ConversionService.cs` 中添加专门的COM异常处理
  - 增加HRESULT错误码显示和用户友好的错误提示
- **改进COM对象管理**: 
  - 修复Visio.Application不支持using语句的问题
  - 添加正确的COM对象清理机制

### 🔍 新增Visio环境检查功能
- **环境检测方法**: 新增 `CheckVisioAvailability()` 方法，支持预先检测Visio环境
- **GUI集成**: 
  - 在主界面新增"🔍 检查Visio"按钮
  - 异步执行环境检查，提供详细的反馈信息
  - 支持在转换前验证Visio是否可用
- **用户体验提升**: 提供步骤化的问题排查指引和解决建议

### 📁 文件变更记录
- 修改: `md2visio/md2visio.csproj` (移除/恢复COM引用配置)
- 修改: `md2visio.GUI/Program.cs` (添加COM线程模式设置)
- 修改: `md2visio/vsdx/@base/VBuilder.cs` (增强错误处理和诊断)
- 修改: `md2visio.GUI/Services/ConversionService.cs` (新增环境检查和异常处理)
- 修改: `md2visio.GUI/Forms/MainForm.cs` (新增检查Visio按钮和相关功能)

### ✅ 修复效果
- **编译成功**: 所有项目正常构建，无编译错误
- **功能完善**: GUI版本现在具备完整的Visio环境检测和错误诊断能力
- **用户友好**: 提供清晰的问题排查步骤和解决建议
- **稳定性提升**: 改善了COM对象的创建、使用和清理流程

---

## 🔧 架构重构 (2025-06-22) - 编译问题修复

### 🏗️ 项目架构调整
- **类库化改造**: 将md2visio项目从可执行文件改为类库 (Library)，解决自包含项目引用冲突
- **编译配置优化**: 
  - 默认配置: md2visio编译为类库，供GUI项目引用
  - ConsoleRelease配置: 保留独立控制台应用编译选项
- **入口点重构**: 删除包含顶级语句的Program.cs，创建ConsoleApp.cs作为控制台入口

### 🛠️ 技术问题解决
- **依赖冲突**: 修复"自包含可执行文件不能由非自包含可执行文件引用"的编译错误
- **模块化设计**: 实现核心功能类库与用户界面分离，提高代码复用性
- **构建成功**: GUI项目编译成功，生成md2visio.GUI.exe可执行文件

### 📁 文件变更记录
- 删除: `md2visio/main/Program.cs` (顶级语句与类库冲突)
- 新增: `md2visio/main/ConsoleApp.cs` (控制台应用入口点)
- 修改: `md2visio/md2visio.csproj` (OutputType: Exe → Library)
- 优化: ConversionService.cs (完善AppConfig调用逻辑)

---

## 🎉 GUI版本发布 (2025-06-22) - 重大更新

### 🆕 全新GUI界面
- **Windows Forms应用**: 创建了完整的图形用户界面版本 (md2visio.GUI)
- **拖拽支持**: 支持拖拽 .md 文件到应用程序进行转换
- **实时反馈**: 转换进度条、实时日志显示、友好的错误提示
- **用户友好**: 黑底绿字终端风格日志、图表类型自动检测、一键打开输出目录

### 🏗️ 项目架构扩展
- **双版本支持**: 同时维护控制台版本和GUI版本，满足不同用户需求
- **代码重用**: ConversionService 封装转换逻辑，实现核心功能的复用
- **代码重用**: 将图表类型检测逻辑从 MainForm 中移至 ConversionService，实现代码复用 
- **事件驱动**: 基于事件的进度和日志更新机制
- **异步处理**: 使用 async/await 避免界面卡顿

### 🎨 界面特性
- **响应式布局**: 使用 TableLayoutPanel 实现自适应布局
- **功能区域划分**: 文件选择、输出设置、转换选项、日志显示等清晰分区
- **状态管理**: 转换中禁用操作按钮，防止重复操作
- **配置保存**: 记住用户的输出目录和选项设置

### 🔧 技术改进
- **AppConfig公开**: 将内部类改为公共类，支持GUI项目调用
- **错误处理**: 完善的异常捕获和用户友好的错误提示
- **路径智能**: GUI版本同样支持模板文件的智能路径查找

### 📖 文档增强
- **GUI使用指南**: 创建详细的GUI版本使用指南和功能介绍
- **技术对比**: 提供命令行版本与GUI版本的优势对比分析
- **扩展建议**: 规划了批量转换、转换历史、主题切换等未来功能

---

## 2025-06-20

*   **模板文件路径问题修复**
    *   修复了 VShapeDrawer 中 md2visio.vssx 模板文件路径问题
    *   实现了智能路径查找机制，支持开发和发布环境
    *   配置了正确的文件复制规则，确保模板文件正确包含在发布包中
    
*   **项目发布配置优化**
    *   配置了单文件自包含发布模式
    *   添加了用户友好的安装助手和使用说明
    *   创建了完整的软件分发包，支持小白用户一键使用

*   **Codelf 知识库初始化**
    *   执行了 `code-inspector` 全量扫描，分析了项目结构、模块和核心调用链。
    *   基于分析报告，通过 `init-codelf` 创建并填充了 `project.md`，完成了知识库的初始构建。