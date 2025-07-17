# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

md2visio是一个将Mermaid图表转换为Microsoft Visio文件的工具，由C# .NET 8开发。项目包含两个主要组件：
- `md2visio`: 核心转换库（.NET类库）
- `md2visio.GUI`: Windows Forms图形界面程序

## 构建和运行命令

### 编译项目
```bash
# 使用Visual Studio 2022打开解决方案文件
# 或使用命令行构建
dotnet build md2visio.sln
```

### 运行GUI程序
```bash
# 在GUI项目目录下
dotnet run --project md2visio.GUI
```

### 运行控制台版本
```bash
# 在核心库目录下，使用ConsoleRelease配置
dotnet run --project md2visio --configuration ConsoleRelease -- /I input.md /O output.vsdx
```

### 发布单文件可执行程序
```bash
# GUI版本
dotnet publish md2visio.GUI -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

# 控制台版本
dotnet publish md2visio -c ConsoleRelease -r win-x64 --self-contained true -p:PublishSingleFile=true
```

## 核心架构

### 处理流程
1. **语法解析**: `md2visio/mermaid/`下的状态机解析器解析Mermaid语法
2. **数据结构转换**: `md2visio/struc/`下定义的抽象语法树(AST)结构
3. **Visio绘制**: `md2visio/vsdx/`下通过COM Interop与Visio交互

### 关键组件
- **FigureBuilderFactory**: `md2visio/struc/figure/FigureBuilderFactory.cs`，负责图表构建的工厂类
- **ConversionService**: `md2visio.GUI/Services/ConversionService.cs`，GUI的转换服务层
- **AppConfig**: `md2visio/main/AppConfig.cs`，应用配置和COM对象生命周期管理
- **状态机解析器**: `md2visio/mermaid/@cmn/`通用组件和各图表类型的解析器

### 支持的图表类型
- ✅ graph/flowchart (流程图)
- ✅ journey (用户旅程图)  
- ✅ pie (饼图)
- ✅ packet-beta (数据包图)
- ✅ xychart-beta (XY图表)
- ❌ 其他类型暂未实现

### 配置文件系统
- `md2visio/default/`: 包含各类图表的默认样式配置YAML文件
- `md2visio/default/theme/`: 主题配置文件
- `md2visio.vssx`: Visio模板文件

## 依赖库
- **Microsoft.Office.Interop.Visio**: Visio COM组件交互
- **YamlDotNet**: YAML配置文件解析
- **System.Drawing.Common**: 图形处理
- **stdole**: COM标准对象库

## COM对象管理

项目重点处理COM对象的生命周期管理：
- AppConfig和ConversionService实现IDisposable模式
- 使用静态AppConfig.Instance管理全局状态
- 显示模式(/V参数)下保持Visio窗口打开
- 非显示模式下自动清理COM对象

## 开发注意事项

### COM组件异常处理
COM异常常见原因：
1. Microsoft Visio未正确安装或注册
2. Visio进程权限不足
3. COM组件损坏

### 多图表支持
- 文件模式：指定具体.vsdx文件名
- 目录模式：自动为多个图表生成带编号的文件

### 线程模型
- GUI程序必须设置STA线程模式(Program.cs:14)
- 异步转换使用Task.Run包装同步方法

## 测试文件
测试Mermaid文件位于`md2visio/test/`目录：
- graph.md
- journey.md  
- packet.md
- pie.md
- xy.md