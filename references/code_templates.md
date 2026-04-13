# VBA 代码模板库

核心模板，按需复制使用。

## 标准Sub结构（现代OOP风格）

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

    ' 性能优化设置（With语句）
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = False
    End With

    ' ===== 主代码区域 =====
    Debug.Print "开始处理: " & wsData.Name
    
    ' 在此处编写主要逻辑（全部使用wsData，绝不碰ActiveSheet）
    ProcessData wsData
    
    ' =====================

CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .StatusBar = False
    End With
    Debug.Print "执行完成，用时: " & Format(Timer - startTime, "0.00") & "秒"
    Exit Sub

ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Private Sub ProcessData(ws As Worksheet)
    ' 纯OOP实现，接收参数不依赖全局状态
    Dim lastRow As Long
    Dim arrData As Variant
    
    ' 获取数据边界
    lastRow = GetLastRow(ws, 1)
    If lastRow < 2 Then Exit Sub
    
    ' 使用With和数组批量处理
    With ws
        arrData = .Range("A2:D" & lastRow).Value
    End With
    
    ' 内存中处理...
    
    ' 写回结果
    With ws
        .Range("A2:D" & lastRow).Value = arrData
    End With
End Sub
```

## 数据边界检测函数

```vba
Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , xlValues, , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function

Function GetLastCol(ws As Worksheet, row As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Rows(row).Find("*", , xlValues, , xlByColumns, xlPrevious)
    GetLastCol = IIf(lastCell Is Nothing, 0, lastCell.Column)
End Function

Function GetUsedRange(ws As Worksheet) As Range
    Dim lastRow As Long, lastCol As Long
    On Error Resume Next
    lastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
    lastCol = ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column
    If lastRow > 0 And lastCol > 0 Then
        Set GetUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    End If
End Function
```

## 数组操作模板

### 基础数组读写（With + Offset/Resize）
```vba
Sub ProcessWithArray(ws As Worksheet)
    Dim wsData As Worksheet
    Dim rngData As Range
    Dim arrData As Variant
    Dim lastRow As Long, rowIdx As Long

    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    
    If lastRow < 2 Then Exit Sub

    ' 使用With和Resize替代字符串拼接
    With wsData
        ' 方法1: Resize
        Set rngData = .Range("A2").Resize(lastRow - 1, 4)
        
        ' 方法2: Offset
        ' Set rngData = .Range("A1:D" & lastRow).Offset(1)
    End With
    
    ' 读取到数组（单次IO）
    arrData = rngData.Value

    ' 内存中处理（使用有意义的下标名）
    For rowIdx = 1 To UBound(arrData, 1)
        arrData(rowIdx, 4) = arrData(rowIdx, 2) * arrData(rowIdx, 3)
    Next rowIdx

    ' 一次性写回（单次IO）
    rngData.Value = arrData
End Sub
```

### 二维数组处理（嵌套With）
```vba
Sub Process2DArray(ws As Worksheet)
    Dim wsData As Worksheet
    Dim arrData As Variant
    Dim rowIdx As Long, colIdx As Long
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    With wsData
        arrData = .Range("A1:D10").Value
    End With
    
    For rowIdx = 1 To UBound(arrData, 1)
        For colIdx = 1 To UBound(arrData, 2)
            ' 处理每个单元格
            arrData(rowIdx, colIdx) = UCase(arrData(rowIdx, colIdx))
        Next colIdx
    Next rowIdx
    
    With wsData
        .Range("A1:D10").Value = arrData
    End With
End Sub
```

### 分块处理大数组
```vba
Sub ProcessLargeArrayInChunks(ws As Worksheet)
    Dim wsData As Worksheet
    Dim arrChunk As Variant
    Dim lastRow As Long
    Dim chunkSize As Long
    Dim startRow As Long, endRow As Long
    Dim rowIdx As Long
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    chunkSize = 5000  ' 每块5000行
    
    For startRow = 2 To lastRow Step chunkSize
        endRow = Application.Min(startRow + chunkSize - 1, lastRow)
        
        With wsData
            Set rngChunk = .Range("A" & startRow & ":D" & endRow)
        End With
        
        arrChunk = rngChunk.Value
        
        ' 处理当前块
        For rowIdx = 1 To UBound(arrChunk, 1)
            arrChunk(rowIdx, 4) = arrChunk(rowIdx, 2) * arrChunk(rowIdx, 3)
        Next rowIdx
        
        rngChunk.Value = arrChunk
        
        ' 释放内存并允许UI更新
        Erase arrChunk
        DoEvents
    Next startRow
End Sub
```

## 文件选择对话框

```vba
Function SelectFile(Optional filter As String = "Excel文件") As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "选择文件"
        .Filters.Clear
        If filter = "Excel文件" Then
            .Filters.Add "Excel文件", "*.xlsx; *.xls; *.xlsm"
        ElseIf filter = "所有文件" Then
            .Filters.Add "所有文件", "*.*"
        Else
            .Filters.Add filter, "*.*"
        End If
        If .Show = -1 Then SelectFile = .SelectedItems(1) Else SelectFile = ""
    End With
End Function

Function SelectFolder() As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "选择文件夹"
        If .Show = -1 Then SelectFolder = .SelectedItems(1) Else SelectFolder = ""
    End With
End Function

Function SelectSaveFile(defaultName As String) As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .AllowMultiSelect = False
        .Title = "保存文件"
        If defaultName <> "" Then
            .InitialFileName = defaultName
        End If
        If .Show = -1 Then SelectSaveFile = .SelectedItems(1) Else SelectSaveFile = ""
    End With
End Function
```

## 进度显示

```vba
Sub ShowProgress(current As Long, total As Long)
    Application.StatusBar = "处理中... " & current & " / " & total & _
                           " (" & Format(current / total, "0%") & ")"
    DoEvents
End Sub

Sub ClearProgress()
    Application.StatusBar = False
End Sub

' 使用示例
Sub ProcessWithProgress()
    Dim i As Long, total As Long
    total = 1000
    
    For i = 1 To total
        ' 处理逻辑
        If i Mod 10 = 0 Then ShowProgress i, total
    Next i
    
    ClearProgress
End Sub
```

## 调试模板

### 基础调试结构
```vba
Sub DebugTemplate()
    Dim startTime As Double
    startTime = Timer

    Debug.Print String(50, "=")
    Debug.Print "开始处理: " & Now
    Debug.Print String(50, "=")

    ' Debug.Assert: 关键检查
    Debug.Assert Not ActiveSheet Is Nothing

    ' Debug.Print: 流程记录
    Debug.Print "处理完成，耗时: " & Format(Timer - startTime, "0.00") & "秒"
End Sub
```

### 变量跟踪
```vba
Sub VariableTracking()
    Dim i As Long, total As Double
    total = 0
    Debug.Print "序号", "当前值", "累计值"
    Debug.Print String(40, "-")
    For i = 1 To 10
        total = total + i
        Debug.Print i, i, total
    Next i
End Sub
```

### 性能计时
```vba
Function StartTimer() As Double
    StartTimer = Timer
    Debug.Print String(50, "=")
    Debug.Print "计时开始: " & Format(Now, "hh:mm:ss")
End Function

Sub EndTimer(startTime As Double, Optional operationName As String = "操作")
    Debug.Print operationName & "耗时: " & Format(Timer - startTime, "0.000") & "秒"
    Debug.Print String(50, "=")
End Sub

' 使用示例
Sub ExampleWithTimer()
    Dim timer As Double
    timer = StartTimer()
    
    ' 处理逻辑...
    
    EndTimer timer, "数据处理"
End Sub
```

## 工作表操作模板

### 创建工作表
```vba
Sub CreateWorksheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
End Sub
```

### 复制工作表
```vba
Sub CopyWorksheet(sourceName As String, newName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sourceName)
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    ActiveSheet.Name = newName
End Sub
```

### 删除工作表（带确认）
```vba
Sub DeleteWorksheet(sheetName As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub
```

---

## For Each遍历模板

**For Each**是遍历对象集合的首选方式，比下标循环更简洁、更安全。

### Range遍历（推荐）
```vba
Sub ProcessRangeWithForEach(ws As Worksheet)
    Dim wsData As Worksheet
    Dim cell As Range          ' 遍历单元格
    Dim rngData As Range
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    ' 定义目标区域
    With wsData
        Set rngData = .Range("A2:A100")
    End With
    
    ' For Each遍历（现代OOP风格）
    For Each cell In rngData
        ' 使用Offset进行相对引用
        If cell.Value > 100 Then
            With cell.Offset(0, 1)
                .Value = "高"
                .Font.Color = RGB(255, 0, 0)
            End With
        End If
    Next cell
End Sub
```

### Worksheet集合遍历
```vba
Sub ProcessAllWorksheets(wb As Workbook)
    Dim ws As Worksheet
    
    ' 遍历工作簿中所有工作表
    For Each ws In wb.Worksheets
        ' 跳过特定名称的工作表
        If ws.Name <> "模板" Then
            ProcessWorksheet ws
        End If
    Next ws
End Sub
```

### Workbook集合遍历
```vba
Sub ProcessAllOpenWorkbooks()
    Dim wb As Workbook
    
    ' 遍历所有打开的工作簿
    For Each wb In Application.Workbooks
        ' 跳过当前工作簿
        If wb.Name <> ThisWorkbook.Name Then
            Debug.Print "处理: " & wb.Name
            ProcessWorkbook wb
        End If
    Next wb
End Sub
```

---

## Dictionary模板

**Dictionary**是VBA中的哈希表，支持O(1)复杂度的键值查找。

### 基础使用模板
```vba
Sub DictionaryBasicExample()
    Dim dict As Object
    Dim key As Variant
    
    ' 创建Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 添加键值对
    dict.Add "苹果", 100
    dict.Add "香蕉", 200
    dict.Add "橙子", 150
    
    ' 检查键是否存在
    If dict.exists("苹果") Then
        Debug.Print "苹果的价格: " & dict("苹果")
    End If
    
    ' 修改值
    dict("苹果") = 120
    
    ' 遍历所有键值对
    For Each key In dict.Keys
        Debug.Print key & ": " & dict(key)
    Next key
    
    ' 删除键
    dict.Remove "香蕉"
    
    ' 清空
    dict.RemoveAll
End Sub
```

### 去重模板（最常用）
```vba
Sub RemoveDuplicatesWithDict(ws As Worksheet)
    Dim wsData As Worksheet
    Dim dict As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim rowIdx As Long
    Dim arrUnique As Variant
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' For Each遍历 + Dictionary去重（O(1)查找）
    With wsData
        For Each cell In .Range("A2:A" & lastRow)
            ' Exists方法检查键是否存在
            If Not dict.exists(cell.Value) Then
                dict.Add cell.Value, cell.Row  ' Key=值, Value=行号
            End If
        Next cell
    End With
    
    ' 获取唯一值数组
    arrUnique = dict.Keys
    
    Debug.Print "唯一值数量: " & dict.Count
    
    ' 输出到新位置
    For rowIdx = 0 To UBound(arrUnique)
        wsData.Cells(rowIdx + 2, "F").Value = arrUnique(rowIdx)
    Next rowIdx
End Sub
```

### 计数统计模板
```vba
Sub CountWithDictionary(ws As Worksheet)
    Dim wsData As Worksheet
    Dim dict As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim key As Variant
    Dim outputRow As Long
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    lastRow = GetLastRow(wsData, 1)
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 统计产品出现次数
    With wsData
        For Each cell In .Range("A2:A" & lastRow)
            If dict.exists(cell.Value) Then
                dict(cell.Value) = dict(cell.Value) + 1  ' 累加
            Else
                dict.Add cell.Value, 1  ' 首次出现
            End If
        Next cell
    End With
    
    ' 输出统计结果
    outputRow = 2
    For Each key In dict.Keys
        With wsData
            .Cells(outputRow, "F").Value = key
            .Cells(outputRow, "G").Value = dict(key)
        End With
        outputRow = outputRow + 1
    Next key
End Sub
```

### Dictionary查找替代数组扫描
```vba
' ❌ 数组扫描（O(n)慢）
Function FindInArray(arr As Variant, target As String) As Long
    Dim i As Long
    For i = 1 To UBound(arr)
        If arr(i) = target Then
            FindInArray = i
            Exit Function
        End If
    Next i
    FindInArray = 0
End Function

' ✅ Dictionary查找（O(1)快）
Sub BuildLookupDictionary()
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIdx As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = ActiveSheet
    lastRow = GetLastRow(ws, 1)
    
    ' 构建查找表
    With ws
        For rowIdx = 2 To lastRow
            ' Key=产品名, Value=行号
            dict.Add .Cells(rowIdx, 1).Value, rowIdx
        Next rowIdx
    End With
    
    ' 查找（O(1)）
    If dict.exists("产品A") Then
        Debug.Print "产品A在第" & dict("产品A") & "行"
    End If
End Sub
```

---

## ListObject模板

**ListObject**（Excel表格）是现代VBA的首选数据结构。

### 基础遍历模板
```vba
Sub ProcessTable(ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim lr As ListRow          ' 表格行对象
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    Set tbl = wsData.ListObjects("数据表")
    
    ' For Each遍历表格行
    For Each lr In tbl.ListRows
        ' 使用列名访问（清晰可读）
        lr.Range(tbl.ListColumns("金额").Index).Value = _
            lr.Range(tbl.ListColumns("数量").Index).Value * _
            lr.Range(tbl.ListColumns("单价").Index).Value
    Next lr
End Sub
```

### 创建表格模板
```vba
Sub CreateTableFromRange(ws As Worksheet)
    Dim wsData As Worksheet
    Dim tbl As ListObject
    Dim rngSource As Range
    
    Set wsData = IIf(ws Is Nothing, ActiveSheet, ws)
    
    ' 获取数据区域
    With wsData
        Set rngSource = .Range("A1").CurrentRegion
    End With
    
    ' 创建表格
    Set tbl = wsData.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rngSource, _
        XlListObjectHasHeaders:=xlYes)
    
    ' 设置属性
    With tbl
        .Name = "销售数据"
        .TableStyle = "TableStyleMedium2"
        .ShowTotals = True
    End With
End Sub
```

### 表格添加行列
```vba
Sub AddDataToTable(ws As Worksheet)
    Dim tbl As ListObject
    Dim lrNew As ListRow
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 添加新行（自动扩展）
    Set lrNew = tbl.ListRows.Add
    
    ' 填充数据（列索引方式，性能更好）
    With lrNew.Range
        .Cells(1, 1).Value = "产品A"      ' 产品名列
        .Cells(1, 2).Value = 10            ' 数量列
        .Cells(1, 3).Value = 100           ' 单价列
        ' 金额列公式自动填充（如果已设置）
    End With
End Sub
```

### 表格筛选和排序
```vba
Sub FilterTable(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 清除现有筛选
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' 应用筛选
    With tbl.Range
        .AutoFilter Field:=2, Criteria1:=">100"     ' 第2列>100
        .AutoFilter Field:=3, Criteria1:="苹果"      ' 第3列=苹果
    End With
End Sub

Sub SortTable(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 清除排序
    tbl.Sort.SortFields.Clear
    
    ' 添加排序字段
    With tbl.Sort.SortFields
        .Add Key:=tbl.ListColumns("金额").Range, Order:=xlDescending
        .Add Key:=tbl.ListColumns("日期").Range, Order:=xlAscending
    End With
    
    ' 应用排序
    With tbl.Sort
        .Header = xlYes
        .Apply
    End With
End Sub
```

### 表格汇总行
```vba
Sub SetupTotalsRow(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    With tbl
        .ShowTotals = True
        
        ' 设置汇总公式
        With .ListColumns("数量")
            .TotalsCalculation = xlTotalsCalculationSum
        End With
        
        With .ListColumns("金额")
            .TotalsCalculation = xlTotalsCalculationSum
            .DataBodyRange.NumberFormat = "#,##0.00"
        End With
        
        ' 平均值
        .ListColumns("单价").TotalsCalculation = xlTotalsCalculationAverage
    End With
End Sub
```

---

## 常用功能函数

### 单元格地址转换
```vba
Function ColLetter(colNum As Long) As String
    Dim temp As Long
    temp = colNum
    ColLetter = ""
    Do While temp > 0
        ColLetter = Chr((temp - 1) Mod 26 + 65) & ColLetter
        temp = (temp - 1) \ 26
    Loop
End Function
```

### 判断工作表是否存在
```vba
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
```

### 判断范围是否为空
```vba
Function IsRangeEmpty(rng As Range) As Boolean
    IsRangeEmpty = Application.WorksheetFunction.CountA(rng) = 0
End Function
```
