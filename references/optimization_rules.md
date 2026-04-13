# VBA 性能优化规则

性能优化黄金法则与实践指南。

## 黄金法则

### 1. 数组替代单元格操作 ⭐最重要

❌ **慢**：逐行读写（N次IO）
```vba
For i = 1 To 10000
    ws.Cells(i, 1).Value = ws.Cells(i, 1).Value * 2
Next i
' IO次数: 20000次（读+写各10000次）
```

✅ **快**：数组批量处理（2次IO）
```vba
data = ws.Range("A1:A10000").Value      ' 1次读
For i = 1 To UBound(data, 1)
    data(i, 1) = data(i, 1) * 2
Next i
ws.Range("A1:A10000").Value = data      ' 1次写
' IO次数: 2次
' 速度提升: 100-1000倍
```

### 2. With语句减少对象调用

❌ **慢**：重复引用对象
```vba
ws.Range("A1").Font.Name = "微软雅黑"
ws.Range("A1").Font.Size = 12
ws.Range("A1").Font.Bold = True
ws.Range("A1").Interior.Color = RGB(200, 200, 200)
' Range对象被解析4次
```

✅ **快**：With语句缓存对象
```vba
With ws.Range("A1")
    .Font.Name = "微软雅黑"
    .Font.Size = 12
    .Font.Bold = True
    .Interior.Color = RGB(200, 200, 200)
End With
' Range对象只解析1次
```

### 3. 关闭屏幕更新和计算

```vba
Sub OptimizeSettings()
    Application.ScreenUpdating = False      ' 禁用屏幕刷新
    Application.Calculation = xlCalculationManual   ' 手动计算
    Application.EnableEvents = False        ' 禁用事件
    
    On Error GoTo CleanUp
    
    ' ===== 主代码 =====
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### 4. Find替代全列扫描

❌ **慢**：End方法
```vba
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
' 如果列中有空单元格，结果可能不准确
```

✅ **快**：Find方法
```vba
Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , , , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function
' 更快、更准确
```

## 循环优化

### 减少循环嵌套
```vba
' 双层循环（慢）
For i = 1 To 1000
    For j = 1 To 1000
        ' 处理
    Next j
Next i

' 单层循环（快）
total = 1000 * 1000
For k = 1 To total
    i = (k - 1) \ 1000 + 1
    j = (k - 1) Mod 1000 + 1
    ' 处理
Next k
```

### 循环内避免不必要计算
```vba
' ❌ 慢：循环内重复计算
For i = 1 To 10000
    result = ws.Cells(i, 1).Value * Application.WorksheetFunction.Sum(ws.Range("Z:Z"))
Next i

' ✅ 快：计算移出循环
sumValue = Application.WorksheetFunction.Sum(ws.Range("Z:Z"))
For i = 1 To 10000
    result = ws.Cells(i, 1).Value * sumValue
Next i
```

### 反向循环（删除时）
```vba
' ✅ 删除行时从后往前
For i = lastRow To 1 Step -1
    If ws.Cells(i, 1).Value = "删除" Then
        ws.Rows(i).Delete
    End If
Next i
```

## 对象优化

### 缓存对象引用
```vba
' ❌ 慢：每次都解析
For i = 1 To 1000
    ThisWorkbook.Worksheets("Sheet1").Cells(i, 1).Value = i
Next i

' ✅ 快：缓存对象
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Sheet1")
For i = 1 To 1000
    ws.Cells(i, 1).Value = i
Next i
```

### 避免Select和Activate
```vba
' ❌ 慢：依赖选择状态
Sheets("Sheet2").Select
Range("A1").Select
Selection.Value = "数据"

' ✅ 快：直接引用
Sheets("Sheet2").Range("A1").Value = "数据"
```

### 及时释放对象
```vba
Sub ProcessWithCleanup()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Open("data.xlsx")
    Set ws = wb.Worksheets(1)
    
    ' 处理逻辑...
    
    ' 及时释放
    Set ws = Nothing
    wb.Close SaveChanges:=False
    Set wb = Nothing
End Sub
```

## 内存优化

### 及时清空数组
```vba
Sub ProcessLargeData()
    Dim largeArray() As Variant
    
    ' 使用大数组
    ReDim largeArray(1 To 100000, 1 To 100)
    ' ... 处理 ...
    
    ' 及时释放内存
    Erase largeArray
End Sub
```

### 分块处理大数据
```vba
Sub ProcessInChunks()
    Dim ws As Worksheet
    Dim data As Variant
    Dim totalRows As Long, chunkSize As Long
    Dim startRow As Long, endRow As Long
    
    Set ws = ActiveSheet
    totalRows = GetLastRow(ws, 1)
    chunkSize = 5000  ' 每次处理5000行
    
    For startRow = 2 To totalRows Step chunkSize
        endRow = Application.Min(startRow + chunkSize - 1, totalRows)
        
        ' 处理当前块
        data = ws.Range("A" & startRow & ":D" & endRow).Value
        ' ... 处理数据 ...
        ws.Range("A" & startRow & ":D" & endRow).Value = data
        
        ' 释放数组
        Erase data
        DoEvents  ' 允许UI更新
    Next startRow
End Sub
```

## 计算优化

### 批量设置公式后转值
```vba
' ❌ 慢：逐个设置公式
For i = 2 To 10000
    ws.Cells(i, 5).Formula = "=B" & i & "*C" & i
Next i

' ✅ 快：批量设置
ws.Range("E2:E10000").Formula = "=B2*C2"
ws.Range("E2:E10000").Value = ws.Range("E2:E10000").Value  ' 转值
```

### 使用数组公式替代循环计算
```vba
' 使用WorksheetFunction替代循环
sumResult = Application.WorksheetFunction.Sum(ws.Range("A:A"))
avgResult = Application.WorksheetFunction.Average(ws.Range("B:B"))
countResult = Application.WorksheetFunction.CountIf(ws.Range("C:C"), ">100")
```

## 优化检查清单

### 必备检查项
- [ ] `Application.ScreenUpdating = False`
- [ ] `Application.Calculation = xlCalculationManual`
- [ ] `Application.EnableEvents = False`
- [ ] 批量操作使用数组替代单元格
- [ ] 使用 `With` 语句减少对象调用
- [ ] 避免 `Select/Activate`
- [ ] 循环内计算移到循环外
- [ ] `Find` 替代整列扫描

### 进阶检查项
- [ ] 对象变量及时释放（`Set obj = Nothing`）
- [ ] 大数组及时清空（`Erase arr`）
- [ ] 大数据分块处理
- [ ] 公式批量设置后转值
- [ ] 使用WorksheetFunction替代简单循环
- [ ] 反向循环处理删除操作

## 性能测试方法

### 基础计时
```vba
Sub MeasurePerformance()
    Dim startTime As Double
    startTime = Timer
    
    ' 要测试的代码
    
    Debug.Print "耗时: " & Format(Timer - startTime, "0.000") & "秒"
End Sub
```

### 对比测试模板
```vba
Sub CompareMethods()
    Dim startTime As Double
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 方法1：慢方法
    startTime = Timer
    ' ... 慢代码 ...
    Debug.Print "方法1耗时: " & Format(Timer - startTime, "0.000") & "秒"
    
    ' 方法2：快方法
    startTime = Timer
    ' ... 快代码 ...
    Debug.Print "方法2耗时: " & Format(Timer - startTime, "0.000") & "秒"
End Sub
```

## 优化效果参考

| 优化项 | 典型提升 |
|-------|---------|
| 数组替代单元格 | 100-1000倍 |
| 关闭ScreenUpdating | 5-10倍 |
| 关闭Calculation | 10-100倍 |
| With语句 | 2-5倍 |
| 避免Select | 2-3倍 |
| Find替代End | 2-5倍 |

