---
name: VBA-vibcoding
description: Excel VBA 智能编程助手 - 代码生成与自动化执行一体化。通过自然语言描述需求，自动生成高质量VBA代码，并直接注入、执行到Excel工作簿中。一站式开发流程：代码生成 → 自动注入 → 执行验证。支持代码模板、性能优化、调试最佳实践、模块管理等完整开发流程。
---

# VBA VibCoding - Excel智能编程助手

一站式VBA开发解决方案：从高质量代码生成到自动化注入执行，让Excel变成可编程的智能工作簿。

## 核心能力

```
┌─────────────────────────────────────────────────────────────┐
│  自然语言需求 → 高质量代码生成 → 自动注入Excel → 执行验证      │
│         ↑                                            ↓      │
│         └──────────── 读取/修改/调试 ←───────────────┘      │
└─────────────────────────────────────────────────────────────┘
```

## 快速开始

### 方式1：完整开发流程（推荐）

```python
# 步骤1: 生成高质量VBA代码（使用本Skill的代码规范）
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
    Debug.Print "处理行数: " & lastRow
    
    If lastRow < 2 Then
        MsgBox "数据不足", vbExclamation
        GoTo CleanUp
    End If
    
    ' 使用数组批量处理
    data = ws.Range("A2:D" & lastRow).Value
    For i = 1 To UBound(data, 1)
        data(i, 4) = data(i, 2) * data(i, 3)
    Next i
    ws.Range("A2:D" & lastRow).Value = data
    
    MsgBox "处理完成! 用时: " & Format(Timer - startTime, "0.00") & "秒", vbInformation
    
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

# 步骤2: 注入到Excel
run_shell_command({
    "command": 'python vba_manager.py write "D:\\data.xlsm" "DataProcessor" "D:\\temp\\data_processor.bas"',
    "dir_path": "{SKILL_DIR}\\tools"
})

# 步骤3: 执行宏
run_shell_command({
    "command": 'python vba_manager.py run "D:\\data.xlsm" "DataProcessor.ProcessData"',
    "dir_path": "{SKILL_DIR}\\tools"
})
```

### 方式2：仅代码生成（手动复制到Excel）

生成代码后，用户可手动在Excel中按 `Alt+F11` 打开VBA编辑器，粘贴代码运行。

## VBA代码生成规范

### 1. 代码质量标准

```vba
' ✅ 必须包含
Option Explicit                    ' 强制变量声明

' ✅ 有意义的变量名
Dim lastRow As Long               ' 好
Dim x As Long                     ' 避免

' ✅ 结构化错误处理
On Error GoTo ErrorHandler
' ... 主代码 ...
CleanUp:
    ' 恢复设置
    Exit Sub
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
```

---

## 现代VBA代码规范（OOP风格）

**目标**：彻底告别「录制宏式」代码，拥抱面向对象编程。

### 10项核心准则

| # | 准则 | 旧式做法 | 现代OOP做法 |
|---|------|---------|------------|
| 1 | **禁止Select/Activate** | `Range("A1").Select: Selection.Copy` | `Range("A1").Copy Destination:=...` |
| 2 | **参数传递对象** | `Sub Macro(): Set ws = ActiveSheet` | `Sub Process(ws As Worksheet)` |
| 3 | **For Each优先** | `For i = 1 To 100: Cells(i,1)...` | `For Each cell In rng...` |
| 4 | **With链式调用** | 重复`ws.Range("A1").Font...` | `With ws.Range("A1"): .Font...` |
| 5 | **ListObject表格** | `Range("A2:D100")` | `tbl.ListObjects("表1").DataBodyRange` |
| 6 | **Dictionary键值对** | 数组+下标查找 | `dict.Exists(key)` 秒查 |
| 7 | **Offset/Resize** | `"A" & row & ":D" & lastRow` | `rng.Offset(1).Resize(n)` |
| 8 | **类型缩写命名** | `Dim sheet1, x, temp` | `Dim wsData, rowIdx, arrCache` |
| 9 | **纯函数设计** | 修改全局状态 | 返回结果，参数明确 |
| 10 | **xlam环境捕获** | 直接使用`ActiveSheet` | `Set wsUser = ActiveSheet`后全用对象 |

### 对比示例：数据处理器

**❌ 旧式录制宏风格**
```vba
Sub OldStyle()
    Dim i As Long
    Sheets("数据").Select           ' 违规1: Select
    Range("A2").Select              ' 违规2: Select
    For i = 1 To 100                ' 违规3: 下标循环
        ActiveCell.Offset(i, 0).Value = i * 2
    Next i
End Sub
```

**✅ 现代OOP风格**
```vba
Sub ModernStyle(ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow               ' 表格行对象
    
    ' 参数为空时安全捕获（仅xlam入口允许）
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    ' 使用ListObject而非裸区域
    Set tbl = wsData.ListObjects("数据表")
    
    ' For Each遍历对象集合
    For Each lr In tbl.ListRows
        ' Offset是相对对象的方法
        lr.Range(1, 4).Value = lr.Range(1, 2).Value * 2
    Next lr
End Sub
```

### 命名规范速查

**类型缩写（强制）**
| 类型 | 缩写 | 示例 |
|------|------|------|
| Worksheet | ws | `wsData`, `wsReport` |
| Workbook | wb | `wbSource`, `wbTarget` |
| Range | rng | `rngInput`, `rngHeader` |
| ListObject | tbl | `tblSales`, `tblData` |
| Dictionary | dict | `dictCache`, `dictLookup` |
| Collection | col | `colItems`, `colResults` |
| Array | arr | `arrData`, `arrBuffer` |

**循环变量命名**
| 场景 | 旧式 | 现代 |
|------|------|------|
| 数组下标 | `i`, `j` | `rowIdx`, `colIdx` |
| 行号 | `i` | `rowNum`, `currentRow` |
| For Each | 无 | `cell`, `lr`, `ws` |

### 2. 性能优化必备

```vba
' 必须设置（主代码前后）
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' 数组替代单元格（关键优化）
' ❌ 慢：逐行读写
For i = 1 To 10000
    ws.Cells(i, 1).Value = ws.Cells(i, 1).Value * 2
Next i

' ✅ 快：数组批量处理
data = ws.Range("A1:A10000").Value
For i = 1 To UBound(data, 1)
    data(i, 1) = data(i, 1) * 2
Next i
ws.Range("A1:A10000").Value = data
```

### 3. 调试优先策略

```
调试优先级: Debug.Print > Debug.Assert > MsgBox
原则: 不中断 > 可控制中断 > 强制中断
```

```vba
' Debug.Print - 流程跟踪（不中断）
Debug.Print "处理第 " & i & " 行，值: " & cellValue

' Debug.Assert - 关键检查（条件False时中断）
Debug.Assert Not ws Is Nothing
Debug.Assert lastRow > 0

' MsgBox - 用户交互（最后手段）
If MsgBox("确定删除？", vbYesNo) = vbYes Then
    ' 执行删除
End If
```

## VBA管理器命令参考

### 命令速查表

| 命令 | 用法 | 说明 |
|------|------|------|
| **list** | `list "<文件路径>"` | 列出所有VBA模块 |
| **read** | `read "<文件路径>" "<模块名>"` | 读取模块代码 |
| **write** | `write "<文件路径>" "<模块名>" "<代码文件>"` | 写入/覆盖模块 |
| **delete** | `delete "<文件路径>" "<模块名>"` | 删除模块 |
| **run** | `run "<文件路径>" "<宏名>" [参数...]` | 运行宏 |
| **export** | `export "<文件路径>" "<模块名>" "<输出文件>"` | 导出模块备份 |
| **import** | `import "<文件路径>" "<模块文件>" [新名称]` | 导入模块 |

### 常用示例

```bash
# 列出现有模块
python vba_manager.py list "D:\\data.xlsm"

# 读取模块代码查看
python vba_manager.py read "D:\\data.xlsm" "Module1"

# 写入新代码（自动创建/覆盖模块）
python vba_manager.py write "D:\\data.xlsm" "MyModule" "D:\\code.bas"

# 运行宏（支持参数）
python vba_manager.py run "D:\\data.xlsm" "MyModule.ProcessData"
python vba_manager.py run "D:\\data.xlsm" "Module1.CalculateSum" "10" "20"

# 导出备份
python vba_manager.py export "D:\\data.xlsm" "Module1" "D:\\backup.bas"

# 导入现有模块
python vba_manager.py import "D:\\data.xlsm" "D:\\library.bas" "Utils"
```

### 返回值处理

所有命令返回JSON格式：

```json
{
  "success": true,
  "module": "MyModule",
  "lines": 45
}
```

```python
import json

result = run_shell_command({...})
data = json.loads(result)

if data["success"]:
    print(f"✅ 成功: {data.get('module', '操作完成')}")
else:
    print(f"❌ 失败: {data['error']}")
```

## 标准代码模板

### 完整Sub结构（现代OOP风格）

```vba
Option Explicit

' 参数传递工作表对象，绝不依赖ActiveSheet
Sub 标准宏(ws As Worksheet)
    Dim wsData As Worksheet
    Dim wsBackup As Worksheet
    Dim startTime As Double
    
    ' xlam安全：明确捕获用户环境后立即Set
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    startTime = Timer
    On Error GoTo ErrorHandler

    ' 自动备份（不使用Activate）
    wsData.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsBackup = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    wsBackup.Name = wsData.Name & "_备份" & Format(Now, "hhmmss")

    ' 性能优化
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With

    ' ===== 主代码区域 =====
    Debug.Print "开始处理: " & wsData.Name
    
    ' 你的代码逻辑（全部使用wsData，绝不碰ActiveSheet）
    ProcessData wsData
    
    ' =====================

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
    End With
    Debug.Print "执行完成，用时: " & Format(Timer - startTime, "0.00") & "秒"
    Exit Sub

ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Private Sub ProcessData(ws As Worksheet)
    ' 纯OOP实现，接收参数不依赖全局状态
End Sub
```

### 数据边界检测函数（With优化）

```vba
Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    With ws.Columns(col)
        Set lastCell = .Find("*", , xlValues, , xlByRows, xlPrevious)
    End With
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function

Function GetLastCol(ws As Worksheet, row As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    With ws.Rows(row)
        Set lastCell = .Find("*", , xlValues, , xlByColumns, xlPrevious)
    End With
    GetLastCol = IIf(lastCell Is Nothing, 0, lastCell.Column)
End Function

Function GetUsedRange(ws As Worksheet) As Range
    Dim lastRow As Long, lastCol As Long
    On Error Resume Next
    With ws.Cells
        lastRow = .Find("*", , , , xlByRows, xlPrevious).Row
        lastCol = .Find("*", , , , xlByColumns, xlPrevious).Column
    End With
    If lastRow > 0 And lastCol > 0 Then
        Set GetUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
End Function
```

### 数组批量操作模板（Offset/Resize）

```vba
Sub ProcessWithArray(ws As Worksheet)
    Dim wsData As Worksheet
    Dim rngData As Range
    Dim arrData As Variant
    Dim lastRow As Long, rowIdx As Long

    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    
    If lastRow < 2 Then Exit Sub

    ' 使用With和Offset/Resize替代字符串拼接
    With wsData
        Set rngData = .Range("A2:D" & lastRow)
        ' 或: Set rngData = .Range("A2").Resize(lastRow - 1, 4)
    End With
    
    arrData = rngData.Value

    ' 内存中处理（使用有意义的下标名）
    For rowIdx = 1 To UBound(arrData, 1)
        arrData(rowIdx, 4) = arrData(rowIdx, 2) * arrData(rowIdx, 3)
    Next rowIdx

    ' 一次性写回
    rngData.Value = arrData
End Sub
```

---

## ListObject操作指南

**ListObject**（Excel表格）是现代VBA的首选数据容器，相比裸区域`Range`具有结构化、自动扩展、公式自动填充等优势。

### ListObject vs Range 对比

| 特性 | Range（裸区域） | ListObject（表格） |
|------|----------------|-------------------|
| 命名引用 | `Range("A2:D100")` | `ListObjects("表1")` |
| 自动扩展 | ❌ 手动调整 | ✅ 自动包含新行 |
| 结构化公式 | ❌ `=A2*B2` | ✅ `=[@数量]*[@单价]` |
| 列引用 | `Range("A:A")` | `ListColumns("姓名")` |
| 遍历 | 下标循环 | `For Each lr In tbl.ListRows` |

### ListObject核心操作

```vba
Sub ProcessTable(ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow          ' 表格行对象
    Dim lc As ListColumn       ' 表格列对象
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    ' 获取表格（通过名称）
    Set tbl = wsData.ListObjects("数据表")
    ' 或: Set tbl = wsData.ListObjects(1)
    
    ' 遍历表格行（For Each - 现代OOP风格）
    For Each lr In tbl.ListRows
        ' 使用列名访问（清晰可读）
        lr.Range(tbl.ListColumns("金额").Index).Value = _
            lr.Range(tbl.ListColumns("数量").Index).Value * _
            lr.Range(tbl.ListColumns("单价").Index).Value
    Next lr
    
    ' 或使用列索引（性能稍好）
    Dim qtyCol As Long, priceCol As Long, amtCol As Long
    qtyCol = tbl.ListColumns("数量").Index
    priceCol = tbl.ListColumns("单价").Index
    amtCol = tbl.ListColumns("金额").Index
    
    For Each lr In tbl.ListRows
        lr.Range(amtCol).Value = lr.Range(qtyCol).Value * lr.Range(priceCol).Value
    Next lr
End Sub
```

### 创建和管理表格

```vba
Sub CreateTable(ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim rngSource As Range
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    ' 定义数据源区域
    With wsData
        Set rngSource = .Range("A1").CurrentRegion
        ' 或: Set rngSource = .Range("A1:D100")
    End With
    
    ' 创建表格
    Set tbl = wsData.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rngSource, _
        XlListObjectHasHeaders:=xlYes)
    
    ' 设置表格名称
    tbl.Name = "销售数据"
    
    ' 设置表格样式
    tbl.TableStyle = "TableStyleMedium2"
    
    ' 显示汇总行
    tbl.ShowTotals = True
    tbl.ListColumns("金额").TotalsCalculation = xlTotalsCalculationSum
End Sub

Sub AddRowsToTable(ws As Worksheet)
    Dim tbl As ListObject
    Dim lrNew As ListRow
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 添加新行（自动扩展）
    Set lrNew = tbl.ListRows.Add
    
    ' 填充数据
    With lrNew.Range
        .Cells(1, 1).Value = "产品A"
        .Cells(1, 2).Value = 10
        .Cells(1, 3).Value = 100
        ' 金额列公式自动填充
    End With
End Sub
```

### 表格常用属性和方法

```vba
Sub TableProperties(ws As Worksheet)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("销售数据")
    
    ' 核心范围
    Debug.Print "Header: " & tbl.HeaderRowRange.Address
    Debug.Print "Data: " & tbl.DataBodyRange.Address
    Debug.Print "Total: " & tbl.TotalsRowRange.Address
    Debug.Print " entire: " & tbl.Range.Address
    
    ' 行列数
    Debug.Print "行数: " & tbl.ListRows.Count
    Debug.Print "列数: " & tbl.ListColumns.Count
    
    ' 列引用
    Dim lc As ListColumn
    For Each lc In tbl.ListColumns
        Debug.Print "列: " & lc.Name & " 索引: " & lc.Index
    Next lc
End Sub
```

---

## Dictionary/Collection使用场景

**Dictionary**（字典）和**Collection**（集合）是VBA中的键值对容器，相比数组具有快速查找、去重、动态扩容等优势。

### 使用场景对比

| 场景 | 数组+下标 | Dictionary | Collection |
|------|----------|------------|------------|
| 去重 | ❌ 需要嵌套循环 O(n²) | ✅ `Exists` 秒查 O(1) | ⚠️ 可但无Key索引 |
| 键值查找 | ❌ 遍历查找 O(n) | ✅ `dict(key)` O(1) | ⚠️ 只能索引访问 |
| 计数统计 | ❌ 复杂逻辑 | ✅ `dict(key) = dict(key) + 1` | ❌ 不支持 |
| 唯一列表 | ❌ Set实现复杂 | ✅ `dict.Keys` 直接返回 | ✅ 天然去重 |

### Dictionary去重模板

```vba
Sub RemoveDuplicatesWithDict(ws As Worksheet)
    Dim wsData As Worksheet
    Dim dict As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim rowIdx As Long
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    
    ' 创建Dictionary对象
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' For Each遍历 + Dictionary去重
    For Each cell In wsData.Range("A2:A" & lastRow)
        ' Exists方法O(1)复杂度查找
        If Not dict.exists(cell.Value) Then
            dict.Add cell.Value, cell.Row  ' Key=值, Value=行号
        End If
    Next cell
    
    ' dict.Keys 即为去重后的值数组
    Dim arrUnique As Variant
    arrUnique = dict.Keys
    
    Debug.Print "唯一值数量: " & dict.Count
    
    ' 输出到新列
    Dim outputRow As Long
    outputRow = 2
    For rowIdx = 0 To UBound(arrUnique)
        wsData.Cells(outputRow, "F").Value = arrUnique(rowIdx)
        outputRow = outputRow + 1
    Next rowIdx
End Sub
```

### Dictionary计数统计

```vba
Sub CountWithDictionary(ws As Worksheet)
    Dim wsData As Worksheet
    Dim dict As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim key As Variant
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 统计产品出现次数
    For Each cell In wsData.Range("A2:A" & lastRow)
        If dict.exists(cell.Value) Then
            dict(cell.Value) = dict(cell.Value) + 1  ' 累加计数
        Else
            dict.Add cell.Value, 1  ' 首次出现
        End If
    Next cell
    
    ' 输出统计结果
    Debug.Print "产品", "次数"
    For Each key In dict.Keys
        Debug.Print key, dict(key)
    Next key
End Sub
```

### Collection使用场景

```vba
Sub CollectionExample()
    Dim col As Collection
    Dim item As Variant
    
    Set col = New Collection
    
    ' 添加元素（自动去重需额外判断）
    col.Add "产品A"
    col.Add "产品B"
    col.Add "产品C"
    
    ' 遍历（For Each）
    For Each item In col
        Debug.Print item
    Next item
    
    ' 索引访问（1-based）
    Debug.Print col(1)  ' 产品A
    
    ' 删除元素
    col.Remove 1
End Sub
```

### Dictionary vs Collection 选择指南

```vba
' 需要快速查找/去重 → 用 Dictionary
Set dict = CreateObject("Scripting.Dictionary")
If dict.exists(key) Then ...

' 只是简单列表/队列 → 用 Collection  
Set col = New Collection
col.Add item
```

---

## xlam加载项专用模板

**xlam**（Excel加载项）代码的特殊要求：
1. 入口点必须捕获用户当前活动环境
2. 捕获后立即赋值给对象变量
3. 后续全部使用对象变量，绝不再碰`ActiveXxx`

### xlam标准结构

```vba
' ============ 模块: MainModule ============

' 公共入口：按钮/快捷键调用
Public Sub XlamMainEntry()
    Dim wbUser As Workbook
    Dim wsUser As Worksheet
    
    ' 1. 捕获用户环境（仅此一处允许ActiveXxx）
    Set wbUser = ActiveWorkbook
    Set wsUser = ActiveSheet
    
    ' 2. 验证捕获成功
    If wbUser Is Nothing Then
        MsgBox "请先打开一个工作簿", vbExclamation
        Exit Sub
    End If
    
    ' 3. 传递对象给处理函数（纯OOP）
    ProcessUserData wbUser, wsUser
End Sub

' 私有处理函数：绝不使用ActiveXxx
Private Sub ProcessUserData(wb As Workbook, ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    
    ' 使用参数传递的对象
    Set wsData = wb.Worksheets("数据")
    Set tbl = wsData.ListObjects("主数据表")
    
    ' 所有操作通过对象变量完成
    ProcessTable tbl
End Sub

Private Sub ProcessTable(tbl As ListObject)
    Dim lr As ListRow
    
    ' 纯OOP：只操作传入的对象
    For Each lr In tbl.ListRows
        ' 处理逻辑
    Next lr
End Sub
```

### xlam多工作簿处理

```vba
Public Sub ProcessMultipleWorkbooks()
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim fd As Office.FileDialog
    Dim filePath As String
    
    ' 选择源文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xlsx; *.xlsm"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With
    
    ' 打开源工作簿（不激活）
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    
    ' 使用当前活动工作簿作为目标（用户环境）
    Set wbTarget = ActiveWorkbook
    
    ' 数据处理
    CopyDataBetweenWorkbooks wbSource, wbTarget
    
    ' 关闭源工作簿（不影响用户Active状态）
    wbSource.Close SaveChanges:=False
End Sub

Private Sub CopyDataBetweenWorkbooks(wbSource As Workbook, wbTarget As Workbook)
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rngSource As Range
    Dim rngTarget As Range
    
    ' 完全不使用Select/Activate
    Set wsSource = wbSource.Worksheets("数据")
    Set wsTarget = wbTarget.Worksheets("汇总")
    
    With wsSource
        Set rngSource = .Range("A1").CurrentRegion
    End With
    
    With wsTarget
        Set rngTarget = .Range("A1")
    End With
    
    ' 直接复制（无剪贴板残留）
    rngSource.Copy Destination:=rngTarget
End Sub
```

### xlam错误处理最佳实践

```vba
Public Sub XlamEntryWithErrorHandling()
    Dim wbUser As Workbook
    Dim wsUser As Worksheet
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    
    ' 捕获原始设置
    With Application
        origScreenUpdating = .ScreenUpdating
        origCalculation = .Calculation
    End With
    
    ' 捕获用户环境
    Set wbUser = ActiveWorkbook
    Set wsUser = ActiveSheet
    
    On Error GoTo ErrorHandler
    
    ' 优化设置
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    ' 主逻辑
    ProcessData wbUser, wsUser
    
CleanUp:
    ' 恢复原始设置（无论成功与否）
    With Application
        .ScreenUpdating = origScreenUpdating
        .Calculation = origCalculation
        .EnableEvents = True
    End With
    Exit Sub
    
ErrorHandler:
    MsgBox "操作失败: " & Err.Description, vbCritical
    Resume CleanUp
End Sub
```

---

## 前置要求

1. **Excel设置**：
   - 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
   - 启用"信任对VBA工程对象模型的访问"

2. **文件格式**：
   - 必须使用 `.xlsm`（启用宏的工作簿）
   - `.xlsx` 无法保存VBA代码

3. **Python依赖**：
   ```bash
   pip install pywin32
   ```

## 知识库导航

详细知识库在 `references/` 目录：

| 文件 | 内容 |
|------|------|
| **code_templates.md** | 完整代码模板库（标准结构、数组操作、文件选择、进度显示） |
| **excel_patterns.md** | Excel自动化模式（单元格、格式、数据、公式、图表、透视表） |
| **optimization_rules.md** | 性能优化规则（数组优化、循环优化、对象优化） |
| **best_practices.md** | 最佳实践清单（命名规范、错误处理、模块化设计） |
| **debugging_best_practices.md** | 调试最佳实践（Debug.Print/Assert/MsgBox使用指南） |

## 触发关键词

**代码生成**: "VBA代码"、"写个宏"、"Excel自动化"、"数据处理脚本"

**执行管理**: "运行宏"、"注入VBA"、"导出模块"、"读取VBA代码"

**调优**: "优化VBA性能"、"调试代码"、"数组处理"

## 最佳实践总结

### 命名规范
| 类型 | 规则 | 示例 |
|-----|------|------|
| 常量 | 全大写下划线 | `MAX_COUNT` |
| 变量 | 驼峰有意义 | `lastRow`, `fileName` |
| 对象 | 类型缩写 | `ws`, `wb`, `rng` |
| 过程 | 动词+名词 | `CalculateSum`, `GetData` |

### 性能检查清单
- [ ] `ScreenUpdating = False`
- [ ] `Calculation = xlCalculationManual`
- [ ] `EnableEvents = False`
- [ ] 批量操作使用数组
- [ ] 使用 `With` 语句
- [ ] 避免 `Select/Activate`
- [ ] 循环内计算移到循环外
- [ ] `Find` 替代整列扫描

### 安全实践
```vba
' 操作前备份
origSheet.Copy After:=Sheets(Sheets.Count)
ActiveSheet.Name = origSheet.Name & "_备份" & Format(Now, "hhmmss")

' 危险操作确认
If MsgBox("确定删除？", vbYesNo + vbQuestion) = vbYes Then
    Application.DisplayAlerts = False
    ' ... 执行操作
    Application.DisplayAlerts = True
End If
```
