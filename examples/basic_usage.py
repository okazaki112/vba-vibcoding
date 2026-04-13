"""
VBA-vibcoding 基础使用示例
展示如何生成VBA代码并注入到Excel中执行
"""

# ============ 示例1: 完整开发流程 ============

def example1_full_workflow():
    """
    完整流程：生成代码 -> 注入Excel -> 执行宏
    """
    # 步骤1: 生成高质量VBA代码
    vba_code = '''Option Explicit

Sub ProcessSalesData()
    Dim ws As Worksheet
    Dim data As Variant
    Dim lastRow As Long, i As Long
    Dim startTime As Double
    Dim totalAmount As Double
    
    On Error GoTo ErrorHandler
    startTime = Timer
    
    ' 性能优化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = ActiveSheet
    lastRow = GetLastRow(ws, 1)
    
    Debug.Print "开始处理销售数据，行数: " & lastRow
    
    If lastRow < 2 Then
        MsgBox "数据不足，请检查数据", vbExclamation
        GoTo CleanUp
    End If
    
    ' 使用数组批量处理
    data = ws.Range("A2:D" & lastRow).Value
    totalAmount = 0
    
    For i = 1 To UBound(data, 1)
        ' 计算金额（数量 * 单价）
        data(i, 4) = data(i, 2) * data(i, 3)
        totalAmount = totalAmount + data(i, 4)
        
        ' 每100行输出进度
        If i Mod 100 = 0 Then
            Debug.Print "已处理: " & i & "/" & UBound(data, 1)
        End If
    Next i
    
    ' 写回结果
    ws.Range("A2:D" & lastRow).Value = data
    
    MsgBox "处理完成!" & vbCrLf & _
           "总行数: " & (lastRow - 1) & vbCrLf & _
           "总金额: " & Format(totalAmount, "#,##0.00") & vbCrLf & _
           "用时: " & Format(Timer - startTime, "0.00") & "秒", _
           vbInformation
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Debug.Print "执行完成，用时: " & Format(Timer - startTime, "0.00") & "秒"
    Exit Sub
    
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , xlValues, , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function
'''
    
    # 写入代码文件
    with open("D:\\temp\\sales_processor.bas", "w", encoding="utf-8") as f:
        f.write(vba_code)
    
    print("✅ 步骤1完成: VBA代码已生成到 D:\\temp\\sales_processor.bas")
    
    # 步骤2: 注入到Excel（使用vba_manager.py）
    # 命令: python vba_manager.py write "D:\data.xlsm" "SalesProcessor" "D:\temp\sales_processor.bas"
    
    # 步骤3: 执行宏
    # 命令: python vba_manager.py run "D:\data.xlsm" "SalesProcessor.ProcessSalesData"


# ============ 示例2: 仅生成代码（手动复制） ============

def example2_generate_only():
    """
    只生成代码，用户可以手动复制到Excel VBA编辑器
    """
    vba_code = '''Option Explicit

Sub FormatReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = GetLastRow(ws, 1)
    
    ' 格式化标题行
    With ws.Range("A1:D1")
        .Font.Name = "微软雅黑"
        .Font.Size = 12
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 添加边框
    With ws.Range("A1:D" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 自动调整列宽
    ws.Columns("A:D").AutoFit
    
    MsgBox "格式化完成!", vbInformation
End Sub

Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , xlValues, , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function
'''
    
    with open("D:\\temp\\format_report.bas", "w", encoding="utf-8") as f:
        f.write(vba_code)
    
    print("✅ 代码已生成，请手动复制到Excel VBA编辑器 (Alt+F11)")
    print(f"文件位置: D:\\temp\\format_report.bas")


# ============ 示例3: VBA管理器命令示例 ============

def example3_manager_commands():
    """
    vba_manager.py 常用命令示例
    """
    commands = """
    # 1. 列出Excel中的所有VBA模块
    python vba_manager.py list "D:\\data.xlsm"
    
    # 2. 读取指定模块的代码
    python vba_manager.py read "D:\\data.xlsm" "Module1"
    
    # 3. 写入/覆盖模块（从.bas文件）
    python vba_manager.py write "D:\\data.xlsm" "MyModule" "D:\\code.bas"
    
    # 4. 运行宏（无参数）
    python vba_manager.py run "D:\\data.xlsm" "MyModule.ProcessData"
    
    # 5. 运行宏（带参数）
    python vba_manager.py run "D:\\data.xlsm" "Module1.Calculate" "10" "20"
    
    # 6. 导出模块备份
    python vba_manager.py export "D:\\data.xlsm" "Module1" "D:\\backup.bas"
    
    # 7. 导入现有模块
    python vba_manager.py import "D:\\data.xlsm" "D:\\library.bas" "Utils"
    
    # 8. 删除模块
    python vba_manager.py delete "D:\\data.xlsm" "OldModule"
    """
    
    print(commands)


# ============ 示例4: 数据处理模板 ============

def example4_data_processing():
    """
    生成一个完整的数据处理模板
    """
    vba_code = '''Option Explicit

' 主处理程序
Sub MainProcess()
    Dim startTime As Double
    startTime = Timer
    
    On Error GoTo ErrorHandler
    
    ' 性能设置
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Debug.Print String(50, "=")
    Debug.Print "开始处理: " & Now
    
    ' 执行各步骤
    Call Step1_ValidateData
    Call Step2_ProcessData
    Call Step3_GenerateReport
    
    Debug.Print "处理完成，总用时: " & Format(Timer - startTime, "0.00") & "秒"
    Debug.Print String(50, "=")
    
    MsgBox "处理完成!", vbInformation
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' 步骤1: 数据验证
Private Sub Step1_ValidateData()
    Debug.Print "[步骤1] 验证数据..."
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Debug.Assert Not ws Is Nothing
    
    Dim lastRow As Long
    lastRow = GetLastRow(ws, 1)
    
    If lastRow < 2 Then
        Err.Raise vbObjectError + 1, , "数据不足"
    End If
    
    Debug.Print "  数据验证通过，行数: " & lastRow
End Sub

' 步骤2: 数据处理
Private Sub Step2_ProcessData()
    Debug.Print "[步骤2] 处理数据..."
    
    Dim ws As Worksheet
    Dim data As Variant
    Dim lastRow As Long, i As Long
    
    Set ws = ActiveSheet
    lastRow = GetLastRow(ws, 1)
    
    ' 使用数组处理
    data = ws.Range("A2:C" & lastRow).Value
    
    For i = 1 To UBound(data, 1)
        ' 示例处理：转换为大写
        data(i, 1) = UCase(data(i, 1))
    Next i
    
    ws.Range("A2:C" & lastRow).Value = data
    
    Debug.Print "  数据处理完成"
End Sub

' 步骤3: 生成报告
Private Sub Step3_GenerateReport()
    Debug.Print "[步骤3] 生成报告..."
    
    ' 格式化
    With ActiveSheet.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    ActiveSheet.Columns("A:C").AutoFit
    
    Debug.Print "  报告生成完成"
End Sub

' 工具函数：获取最后行号
Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , xlValues, , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function
'''
    
    with open("D:\\temp\\main_process.bas", "w", encoding="utf-8") as f:
        f.write(vba_code)
    
    print("✅ 数据处理模板已生成: D:\\temp\\main_process.bas")


if __name__ == "__main__":
    print("="*60)
    print("VBA-vibcoding 使用示例")
    print("="*60)
    print()
    
    print("示例1: 完整开发流程")
    example1_full_workflow()
    print()
    
    print("示例2: 仅生成代码")
    example2_generate_only()
    print()
    
    print("示例3: VBA管理器命令")
    example3_manager_commands()
    print()
    
    print("示例4: 数据处理模板")
    example4_data_processing()
