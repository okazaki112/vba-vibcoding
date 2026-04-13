# VBA VibCoding 🚀

> Excel VBA 智能编程助手 - 从自然语言需求到自动化代码生成与执行

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Excel](https://img.shields.io/badge/Excel-2016+-green.svg)](https://www.microsoft.com/excel)

## 项目简介

VBA VibCoding 是一个一站式 VBA 开发解决方案，通过自然语言描述需求，自动生成高质量 VBA 代码，并直接注入、执行到 Excel 工作簿中。让 Excel 变成可编程的智能工作簿！

### 核心能力

```
┌─────────────────────────────────────────────────────────────┐
│  自然语言需求 → 高质量代码生成 → 自动注入Excel → 执行验证      │
│         ↑                                            ↓      │
│         └──────────── 读取/修改/调试 ←───────────────┘      │
└─────────────────────────────────────────────────────────────┘
```

## ✨ 主要特性

- 🤖 **AI 驱动代码生成** - 自然语言描述即可生成高质量 VBA 代码
- ⚡ **自动化注入执行** - 一键将代码注入 Excel 并执行
- 📚 **现代 OOP 规范** - 告别录制宏风格，拥抱面向对象编程
- 🔧 **完整开发工具链** - 代码生成、注入、调试、导出全流程支持
- 📖 **丰富代码模板** - ListObject、Dictionary、数组优化等现代 VBA 模式
- 🛡️ **安全错误处理** - 结构化错误处理与性能优化最佳实践

## 🚀 快速开始

### 前置要求

1. **Python 环境**
   ```bash
   pip install pywin32
   ```

2. **Excel 设置**
   - 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
   - 启用"信任对 VBA 工程对象模型的访问"

3. **文件格式**
   - 必须使用 `.xlsm`（启用宏的工作簿）

### 方式 1：完整开发流程（推荐）

```python
# 步骤1: 生成高质量 VBA 代码
write_file({
    "file_path": "D:\\temp\\data_processor.bas",
    "content": '''Option Explicit

Sub ProcessData()
    Dim ws As Worksheet
    Dim data As Variant
    Dim lastRow As Long, i As Long
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    startTime = Timer
    
    ' 性能优化设置
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = ActiveSheet
    lastRow = GetLastRow(ws, 1)
    
    ' 使用数组批量处理
    data = ws.Range("A2:D" & lastRow).Value
    For i = 1 To UBound(data, 1)
        data(i, 4) = data(i, 2) * data(i, 3)
    Next i
    ws.Range("A2:D" & lastRow).Value = data
    
    MsgBox "处理完成! 用时: " & Format(Timer - startTime, "0.00") & "秒"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , , , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function'''
})

# 步骤2: 注入到 Excel
python vba_manager.py write "D:\\data.xlsm" "DataProcessor" "D:\\temp\\data_processor.bas"

# 步骤3: 执行宏
python vba_manager.py run "D:\\data.xlsm" "DataProcessor.ProcessData"
```

### 方式 2：手动复制代码

生成代码后，在 Excel 中按 `Alt+F11` 打开 VBA 编辑器，粘贴代码即可运行。

## 🛠️ VBA 管理器工具

`vba_manager.py` 提供完整的 VBA 模块管理功能：

| 命令 | 用法 | 说明 |
|------|------|------|
| **list** | `list "<文件路径>"` | 列出所有 VBA 模块 |
| **read** | `read "<文件路径>" "<模块名>"` | 读取模块代码 |
| **write** | `write "<文件路径>" "<模块名>" "<代码文件>"` | 写入/覆盖模块 |
| **delete** | `delete "<文件路径>" "<模块名>"` | 删除模块 |
| **run** | `run "<文件路径>" "<宏名>" [参数...]` | 运行宏 |
| **export** | `export "<文件路径>" "<模块名>" "<输出文件>"` | 导出模块备份 |
| **import** | `import "<文件路径>" "<模块文件>" [新名称]` | 导入模块 |

### 使用示例

```bash
# 列出现有模块
python vba_manager.py list "D:\\data.xlsm"

# 写入新代码
python vba_manager.py write "D:\\data.xlsm" "MyModule" "code.bas"

# 运行宏
python vba_manager.py run "D:\\data.xlsm" "MyModule.ProcessData"

# 导出备份
python vba_manager.py export "D:\\data.xlsm" "Module1" "backup.bas"
```

## 📚 现代 VBA 编程规范

### 10 项核心准则

| # | 准则 | 旧式做法 | 现代 OOP 做法 |
|---|------|---------|--------------|
| 1 | **禁止 Select/Activate** | `Range("A1").Select: Selection.Copy` | `Range("A1").Copy Destination:=...` |
| 2 | **参数传递对象** | `Sub Macro(): Set ws = ActiveSheet` | `Sub Process(ws As Worksheet)` |
| 3 | **For Each 优先** | `For i = 1 To 100: Cells(i,1)...` | `For Each cell In rng...` |
| 4 | **With 链式调用** | 重复 `ws.Range("A1").Font...` | `With ws.Range("A1"): .Font...` |
| 5 | **ListObject 表格** | `Range("A2:D100")` | `tbl.ListObjects("表1").DataBodyRange` |
| 6 | **Dictionary 键值对** | 数组+下标查找 | `dict.Exists(key)` 秒查 |
| 7 | **Offset/Resize** | `"A" & row & ":D" & lastRow` | `rng.Offset(1).Resize(n)` |
| 8 | **类型缩写命名** | `Dim sheet1, x, temp` | `Dim wsData, rowIdx, arrCache` |
| 9 | **纯函数设计** | 修改全局状态 | 返回结果，参数明确 |
| 10 | **xlam 环境捕获** | 直接使用 `ActiveSheet` | `Set wsUser = ActiveSheet` 后全用对象 |

### 代码对比示例

**❌ 旧式录制宏风格**
```vba
Sub OldStyle()
    Sheets("数据").Select
    Range("A2").Select
    For i = 1 To 100
        ActiveCell.Offset(i, 0).Value = i * 2
    Next i
End Sub
```

**✅ 现代 OOP 风格**
```vba
Sub ModernStyle(ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    Set tbl = wsData.ListObjects("数据表")
    
    For Each lr In tbl.ListRows
        lr.Range(1, 4).Value = lr.Range(1, 2).Value * 2
    Next lr
End Sub
```

## 📖 文档导航

| 文档 | 内容 |
|------|------|
| [SKILL.md](SKILL.md) | 完整技能文档与 API 参考 |
| [docs/API.md](docs/API.md) | API 详细文档 |
| [references/code_templates.md](references/code_templates.md) | 代码模板库 |
| [references/excel_patterns.md](references/excel_patterns.md) | Excel 自动化模式 |
| [references/best_practices.md](references/best_practices.md) | 最佳实践清单 |
| [references/optimization_rules.md](references/optimization_rules.md) | 性能优化规则 |
| [references/debugging_best_practices.md](references/debugging_best_practices.md) | 调试指南 |
| [examples/basic_usage.py](examples/basic_usage.py) | 基础使用示例 |

## 🤝 贡献指南

欢迎提交 Issue 和 Pull Request！请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解详情。

## 📄 许可证

本项目采用 [MIT 许可证](LICENSE) - 详见 LICENSE 文件。

## 🙏 致谢

感谢所有为本项目做出贡献的开发者！

## 📮 联系方式

- 项目主页：[https://github.com/yourusername/vba-vibcoding](https://github.com/yourusername/vba-vibcoding)
- 问题反馈：[Issues](https://github.com/yourusername/vba-vibcoding/issues)

---

⭐ **如果这个项目对你有帮助，请给它一个 Star！**
