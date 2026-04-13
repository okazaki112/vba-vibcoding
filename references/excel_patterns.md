# Excel 自动化模式

常用Excel操作的核心模式。

## 单元格操作

### 读取单元格
```vba
' 基础读取
value = ws.Range("A1").Value
value = ws.Cells(1, 1).Value
value = ws.Cells(1, "A").Value

' 安全读取（处理错误值）
If Not IsError(ws.Range("A1").Value) Then
    value = ws.Range("A1").Value
Else
    value = 0
End If

' 读取公式
formula = ws.Range("A1").Formula
formulaR1C1 = ws.Range("A1").FormulaR1C1
```

### 写入单元格
```vba
' 基础写入
ws.Range("A1").Value = "数据"
ws.Cells(1, 1).Value = 100

' 批量写入（数组）
ws.Range("A1:D10").Value = dataArray

' 填充公式
ws.Range("E2:E100").Formula = "=B2*C2"
ws.Range("E2:E100").FormulaR1C1 = "=RC[-3]*RC[-2]"

' 公式转值
ws.Range("E2:E100").Value = ws.Range("E2:E100").Value
```

### 复制粘贴
```vba
' 直接复制到目标
ws.Range("A1:D10").Copy Destination:=ws.Range("F1")

' 使用剪贴板
ws.Range("A1:D10").Copy
ws.Range("F1").PasteSpecial xlPasteValues
ws.Range("F1").PasteSpecial xlPasteFormulas
ws.Range("F1").PasteSpecial xlPasteFormats
Application.CutCopyMode = False
```

## 格式化

### 字体和颜色
```vba
With ws.Range("A1:D10")
    ' 字体设置
    .Font.Name = "微软雅黑"
    .Font.Size = 11
    .Font.Bold = True
    .Font.Italic = False
    .Font.Color = RGB(0, 0, 0)
    
    ' 背景色
    .Interior.Color = RGB(200, 200, 200)
    .Interior.Pattern = xlSolid
    
    ' 对齐
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = True
    
    ' 边框
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    .Borders.Color = RGB(0, 0, 0)
    
    ' 数字格式
    .NumberFormat = "0.00"
    .NumberFormat = "yyyy-mm-dd"
    .NumberFormat = "#,##0"
End With
```

### 条件格式
```vba
' 单元格值条件
With ws.Range("A2:A100").FormatConditions.Add(xlCellValue, xlGreater, 100)
    .Interior.Color = RGB(255, 200, 200)
    .Font.Color = RGB(255, 0, 0)
End With

' 公式条件
With ws.Range("A2:A100").FormatConditions.Add(xlExpression, , "=B2>100")
    .Interior.Color = RGB(200, 255, 200)
End With

' 清除条件格式
ws.Range("A2:A100").FormatConditions.Delete
```

### 行列格式
```vba
' 自动调整
ws.Columns("A:D").AutoFit
ws.Rows("1:10").AutoFit

' 固定宽度
ws.Columns("A").ColumnWidth = 15
ws.Rows("1").RowHeight = 30

' 隐藏/显示
ws.Columns("A").Hidden = True
ws.Rows("1").Hidden = False
```

## 数据处理

### 筛选
```vba
' 添加筛选
ws.Range("A1").AutoFilter

' 按值筛选
ws.Range("A1").AutoFilter Field:=1, Criteria1:="苹果"

' 多条件筛选
ws.Range("A1").AutoFilter Field:=1, Criteria1:=Array("苹果", "香蕉"), Operator:=xlFilterValues

' 清除筛选
ws.Range("A1").AutoFilter Field:=1

' 关闭筛选
ws.Range("A1").AutoFilter Mode:=False
```

### 排序
```vba
With ws.Sort
    .SortFields.Clear
    .SortFields.Add Key:=ws.Range("A2"), Order:=xlAscending
    .SortFields.Add Key:=ws.Range("B2"), Order:=xlDescending
    .SetRange ws.Range("A1:D100")
    .Header = xlYes
    .Apply
End With
```

### 删除重复项
```vba
ws.Range("A1:D100").RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
```

### 查找替换
```vba
' 查找
Dim foundCell As Range
Set foundCell = ws.Range("A:A").Find(What:="苹果", LookIn:=xlValues)
If Not foundCell Is Nothing Then
    foundCell.Value = "香蕉"
End If

' 查找全部
Dim foundRange As Range
Set foundCell = ws.Range("A:A").Find("苹果")
If Not foundCell Is Nothing Then
    Set foundRange = foundCell
    Set foundCell = ws.Range("A:A").FindNext(foundCell)
    Do While foundCell.Address <> foundRange.Address
        Set foundRange = Union(foundRange, foundCell)
        Set foundCell = ws.Range("A:A").FindNext(foundCell)
    Loop
End If

' 替换
ws.Range("A:A").Replace What:="旧值", Replacement:="新值"
```

## 图表操作

### 创建图表
```vba
Sub CreateChart()
    Dim cht As Chart
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 创建图表
    Set cht = Charts.Add
    With cht
        .SetSourceData Source:=ws.Range("A1:B10")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "销售数据"
        .Location Where:=xlLocationAsObject, Name:=ws.Name
    End With
End Sub
```

### 图表类型常量
- `xlColumnClustered` - 簇状柱形图
- `xlColumnStacked` - 堆积柱形图
- `xlLine` - 折线图
- `xlLineMarkers` - 带数据标记的折线图
- `xlPie` - 饼图
- `xlBarClustered` - 簇状条形图
- `xlArea` - 面积图
- `xlXYScatter` - 散点图

### 图表格式化
```vba
With cht
    ' 标题
    .HasTitle = True
    .ChartTitle.Text = "图表标题"
    .ChartTitle.Font.Size = 14
    .ChartTitle.Font.Bold = True
    
    ' 图例
    .HasLegend = True
    .Legend.Position = xlLegendPositionBottom
    
    ' 坐标轴
    .Axes(xlCategory).HasTitle = True
    .Axes(xlCategory).AxisTitle.Text = "X轴"
    .Axes(xlValue).HasTitle = True
    .Axes(xlValue).AxisTitle.Text = "Y轴"
End With
```

## 数据透视表

### 创建透视表
```vba
Sub CreatePivotTable()
    Dim ptCache As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    Set dataRange = ws.Range("A1").CurrentRegion
    
    ' 创建缓存
    Set ptCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    ' 创建透视表
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=ws.Cells(2, 10), _
        TableName:="透视表1")
    
    ' 设置字段
    With pt
        .PivotFields("类别").Orientation = xlRowField
        .PivotFields("月份").Orientation = xlColumnField
        .AddDataField .PivotFields("金额"), "金额求和", xlSum
        .AddDataField .PivotFields("数量"), "数量求和", xlSum
    End With
End Sub
```

### 透视表字段位置
- `xlRowField` - 行区域
- `xlColumnField` - 列区域
- `xlDataField` - 数据区域
- `xlPageField` - 筛选区域

### 聚合函数
- `xlSum` - 求和
- `xlCount` - 计数
- `xlAverage` - 平均值
- `xlMax` - 最大值
- `xlMin` - 最小值
- `xlProduct` - 乘积

## 工作表事件

### Worksheet_Change（单元格变化）
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 只在A列变化时触发
    If Not Intersect(Target, Me.Range("A:A")) Is Nothing Then
        Application.EnableEvents = False
        Target.Offset(0, 1).Value = Now
        Application.EnableEvents = True
    End If
End Sub
```

### Worksheet_SelectionChange（选择变化）
```vba
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' 高亮当前行
    Cells.Interior.ColorIndex = xlNone
    Target.EntireRow.Interior.Color = RGB(240, 240, 240)
End Sub
```

### Worksheet_BeforeDoubleClick（双击）
```vba
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column = 1 Then
        MsgBox "你双击了A列"
        Cancel = True
    End If
End Sub
```

## 工作簿事件

```vba
' 必须放在ThisWorkbook对象中

Private Sub Workbook_Open()
    MsgBox "欢迎使用！"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If MsgBox("确定要关闭吗？", vbYesNo) = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Debug.Print "工作表 " & Sh.Name & " 发生变化"
End Sub
```

---

## ListObject结构化引用

**ListObject**（Excel表格）提供了比Range更结构化的数据操作方式。

### 表格vs区域对比

| 操作 | Range（裸区域） | ListObject（表格） |
|------|----------------|-------------------|
| 引用数据 | `Range("A2:D100")` | `tbl.DataBodyRange` |
| 引用标题 | `Range("A1:D1")` | `tbl.HeaderRowRange` |
| 添加行 | 手动调整Range | `tbl.ListRows.Add` |
| 遍历数据 | `For i = 2 To 100` | `For Each lr In tbl.ListRows` |
| 列引用 | `Range("B:B")` | `tbl.ListColumns("姓名")` |
| 汇总公式 | 手动输入 | `tbl.ShowTotals = True` |

### 核心范围和属性

```vba
Sub ListObjectProperties(ws As Worksheet)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("销售数据")
    
    ' 核心范围
    Debug.Print "整个表格: " & tbl.Range.Address
    Debug.Print "标题行: " & tbl.HeaderRowRange.Address
    Debug.Print "数据区域: " & tbl.DataBodyRange.Address
    Debug.Print "汇总行: " & tbl.TotalsRowRange.Address
    
    ' 行列数
    Debug.Print "总行数: " & tbl.ListRows.Count
    Debug.Print "总列数: " & tbl.ListColumns.Count
    
    ' 表格属性
    Debug.Print "表格名称: " & tbl.Name
    Debug.Print "显示汇总: " & tbl.ShowTotals
    Debug.Print "表格样式: " & tbl.TableStyle
End Sub
```

### 列操作

```vba
Sub ColumnOperations(ws As Worksheet)
    Dim tbl As ListObject
    Dim lc As ListColumn
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 通过名称获取列
    Set lc = tbl.ListColumns("金额")
    Debug.Print "金额列索引: " & lc.Index
    Debug.Print "金额列范围: " & lc.Range.Address
    Debug.Print "金额数据: " & lc.DataBodyRange.Address
    
    ' 通过索引获取列
    Set lc = tbl.ListColumns(2)
    Debug.Print "第2列名称: " & lc.Name
    
    ' 遍历所有列
    For Each lc In tbl.ListColumns
        Debug.Print "列名: " & lc.Name & ", 索引: " & lc.Index
    Next lc
End Sub
```

### 行操作

```vba
Sub RowOperations(ws As Worksheet)
    Dim tbl As ListObject
    Dim lr As ListRow
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 添加新行
    Set lr = tbl.ListRows.Add
    lr.Range.Cells(1, 1).Value = "新产品"
    
    ' 插入到指定位置
    Set lr = tbl.ListRows.Add(2)  ' 插入到第2行
    
    ' 删除行
    tbl.ListRows(5).Delete
    
    ' 遍历所有行
    For Each lr In tbl.ListRows
        Debug.Print "第" & lr.Index & "行数据"
    Next lr
End Sub
```

### 单元格访问

```vba
Sub CellAccess(ws As Worksheet)
    Dim tbl As ListObject
    Dim lr As ListRow
    Dim qtyCol As Long, priceCol As Long
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 获取列索引（只需查找一次）
    qtyCol = tbl.ListColumns("数量").Index
    priceCol = tbl.ListColumns("单价").Index
    
    ' 遍历行并访问单元格
    For Each lr In tbl.ListRows
        ' 方法1: 通过列名（清晰但稍慢）
        lr.Range(tbl.ListColumns("金额").Index).Value = _
            lr.Range(tbl.ListColumns("数量").Index).Value * _
            lr.Range(tbl.ListColumns("单价").Index).Value
        
        ' 方法2: 通过预存索引（性能更好）
        lr.Range(qtyCol + 2).Value = _
            lr.Range(qtyCol).Value * lr.Range(priceCol).Value
    Next lr
End Sub
```

### 结构化引用公式

```vba
Sub StructuredReferences(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 设置公式（使用结构化引用）
    With tbl.ListColumns("金额").DataBodyRange
        .Formula = "=[@数量]*[@单价]"
        ' 或R1C1风格
        .FormulaR1C1 = "=RC[-2]*RC[-1]"
    End With
    
    ' 公式转值
    With tbl.ListColumns("金额").DataBodyRange
        .Value = .Value
    End With
End Sub
```

### 表格筛选

```vba
Sub FilterTable(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 清除筛选
    On Error Resume Next
    tbl.AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' 单列筛选
    tbl.Range.AutoFilter Field:=2, Criteria1:="苹果"
    
    ' 多条件筛选
    tbl.Range.AutoFilter Field:=3, Criteria1:=">100"
    
    ' 多值筛选
    tbl.Range.AutoFilter Field:=1, Criteria1:=Array("苹果", "香蕉"), Operator:=xlFilterValues
End Sub
```

### 表格排序

```vba
Sub SortTable(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    ' 清除排序
    tbl.Sort.SortFields.Clear
    
    ' 添加排序条件
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

### 汇总行

```vba
Sub TotalsRow(ws As Worksheet)
    Dim tbl As ListObject
    
    Set tbl = ws.ListObjects("销售数据")
    
    With tbl
        ' 显示汇总行
        .ShowTotals = True
        
        ' 求和
        .ListColumns("数量").TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("金额").TotalsCalculation = xlTotalsCalculationSum
        
        ' 平均值
        .ListColumns("单价").TotalsCalculation = xlTotalsCalculationAverage
        
        ' 计数
        .ListColumns("产品").TotalsCalculation = xlTotalsCalculationCount
        
        ' 自定义公式
        .ListColumns("备注").TotalsRowRange.Formula = "=COUNTIF([金额],">1000")"
    End With
End Sub
```

### 创建和删除表格

```vba
Sub CreateAndDeleteTable(ws As Worksheet)
    Dim tbl As ListObject
    Dim rngData As Range
    
    ' 创建表格
    With ws
        Set rngData = .Range("A1").CurrentRegion
    End With
    
    Set tbl = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=rngData, _
        XlListObjectHasHeaders:=xlYes)
    
    With tbl
        .Name = "新表格"
        .TableStyle = "TableStyleMedium2"
    End With
    
    ' 删除表格（保留数据）
    tbl.Unlink
    
    ' 彻底删除（数据和表格）
    ' tbl.Delete
End Sub
```

### 表格样式

```vba
Sub TableStyles(ws As Worksheet)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("销售数据")
    
    ' 常用样式
    tbl.TableStyle = "TableStyleLight1"    ' 浅色
    tbl.TableStyle = "TableStyleMedium2"   ' 中等
    tbl.TableStyle = "TableStyleDark1"     ' 深色
    
    ' 无样式
    tbl.TableStyle = ""
End Sub
```
