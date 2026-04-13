# VBA-vibcoding API文档

## 概述

VBA-vibcoding 提供两个层面的功能：
1. **代码生成**：基于最佳实践生成高质量VBA代码
2. **执行管理**：通过Python工具与Excel交互，实现代码注入和执行

## 核心组件

### 1. SKILL.md
主技能文档，包含：
- 快速开始指南
- 代码生成规范
- VBA管理器命令参考
- 标准代码模板

### 2. vba_manager.py
Excel VBA管理工具，提供命令行接口。

### 3. references/
详细知识库，包含5个参考文档。

## vba_manager.py 详细API

### 命令列表

#### list - 列出模块
```bash
python vba_manager.py list "<Excel文件路径>"
```

**返回**:
```json
{
  "success": true,
  "file": "data.xlsm",
  "modules": [
    {"name": "Module1", "type": "模块", "type_id": 1, "lines": 45},
    {"name": "Sheet1", "type": "文档", "type_id": 100, "lines": 10}
  ]
}
```

#### read - 读取模块代码
```bash
python vba_manager.py read "<Excel文件路径>" "<模块名>"
```

**返回**:
```json
{
  "success": true,
  "module": "Module1",
  "code": "Sub Hello()\n    MsgBox \"Hello\"\nEnd Sub"
}
```

#### write - 写入/覆盖模块
```bash
python vba_manager.py write "<Excel文件路径>" "<模块名>" "<代码文件路径>"
```

**返回**:
```json
{
  "success": true,
  "module": "MyModule",
  "lines": 45
}
```

**说明**:
- 如果模块已存在，会先删除再创建
- 代码文件必须是UTF-8编码的.bas文件
- 操作完成后自动保存Excel

#### delete - 删除模块
```bash
python vba_manager.py delete "<Excel文件路径>" "<模块名>"
```

**返回**:
```json
{
  "success": true,
  "deleted": "OldModule"
}
```

#### run - 运行宏
```bash
python vba_manager.py run "<Excel文件路径>" "<宏名>" [参数1] [参数2] ...
```

**返回**:
```json
{
  "success": true,
  "macro": "Module1.ProcessData",
  "result": "返回值（如果有）"
}
```

**宏名格式**:
- `ModuleName.SubName` - 指定模块的子程序
- `SubName` - 默认模块的子程序

#### export - 导出模块
```bash
python vba_manager.py export "<Excel文件路径>" "<模块名>" "<输出文件路径>"
```

**返回**:
```json
{
  "success": true,
  "module": "Module1",
  "output": "D:\\backup.bas"
}
```

#### import - 导入模块
```bash
python vba_manager.py import "<Excel文件路径>" "<模块文件路径>" [新模块名]
```

**返回**:
```json
{
  "success": true,
  "imported": "Utils"
}
```

## Python调用示例

### 基础调用
```python
import subprocess
import json

# 列出模块
result = subprocess.run(
    ['python', 'vba_manager.py', 'list', 'D:\\data.xlsm'],
    capture_output=True,
    text=True,
    cwd='{SKILL_DIR}\\tools'
)

data = json.loads(result.stdout)
if data['success']:
    for module in data['modules']:
        print(f"{module['name']}: {module['lines']}行")
```

### 完整工作流
```python
import subprocess
import json

def vba_manager_command(command, *args):
    """调用vba_manager.py的封装函数"""
    cmd = ['python', 'vba_manager.py', command] + list(args)
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd='{SKILL_DIR}\\tools'
    )
    return json.loads(result.stdout)

# 1. 写入代码
result = vba_manager_command(
    'write',
    'D:\\data.xlsm',
    'MyModule',
    'D:\\code.bas'
)
print(f"写入结果: {result}")

# 2. 运行宏
result = vba_manager_command(
    'run',
    'D:\\data.xlsm',
    'MyModule.ProcessData'
)
print(f"运行结果: {result}")

# 3. 导出备份
result = vba_manager_command(
    'export',
    'D:\\data.xlsm',
    'MyModule',
    'D:\\backup.bas'
)
print(f"导出结果: {result}")
```

## 代码生成模板API

### 标准Sub结构
```vba
Sub [名称]()
    Dim origSheet As Worksheet
    Dim startTime As Double
    startTime = Timer
    On Error GoTo ErrorHandler

    ' 自动备份
    Set origSheet = ActiveSheet
    origSheet.Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = origSheet.Name & "_备份" & Format(Now, "hhmmss")
    origSheet.Activate

    ' 性能优化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ===== 主代码 =====
    
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
```

### 数据边界检测函数
```vba
Function GetLastRow(ws As Worksheet, col As Long) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Columns(col).Find("*", , xlValues, , xlByRows, xlPrevious)
    GetLastRow = IIf(lastCell Is Nothing, 0, lastCell.Row)
End Function
```

### 数组处理模板
```vba
' 读取数据到数组
data = ws.Range("A2:D" & lastRow).Value

' 内存中处理（快速）
For i = 1 To UBound(data, 1)
    data(i, 4) = data(i, 2) * data(i, 3)
Next i

' 写回工作表
ws.Range("A2:D" & lastRow).Value = data
```

## 错误处理

### 常见错误码

| 错误 | 原因 | 解决方案 |
|-----|------|---------|
| 文件不存在 | Excel文件路径错误 | 检查文件路径 |
| 模块不存在 | 模块名错误 | 使用list命令查看可用模块 |
| 宏未找到 | 宏名错误或模块未编译 | 检查宏名，确保代码无语法错误 |
| 权限拒绝 | Excel未启用VBA访问 | 启用"信任对VBA工程对象模型的访问" |
| 类型不匹配 | 参数类型错误 | 检查传递给宏的参数类型 |

### 错误处理示例
```python
import subprocess
import json

result = subprocess.run(
    ['python', 'vba_manager.py', 'run', 'D:\\data.xlsm', 'InvalidMacro'],
    capture_output=True,
    text=True
)

try:
    data = json.loads(result.stdout)
    if not data['success']:
        print(f"❌ 错误: {data['error']}")
        # 根据错误类型处理
        if '不存在' in data['error']:
            print("提示: 使用 'list' 命令查看可用模块")
        elif '权限' in data['error']:
            print("提示: 检查Excel宏设置")
except json.JSONDecodeError:
    print(f"❌ 命令执行失败: {result.stderr}")
```

## 前置要求

### Excel设置
1. 文件 → 选项 → 信任中心 → 信任中心设置
2. 宏设置 → 启用"信任对VBA工程对象模型的访问"

### Python依赖
```bash
pip install pywin32
```

### 文件格式
- 必须使用 `.xlsm`（启用宏的工作簿）
- `.xlsx` 无法保存VBA代码

## 性能建议

### 大批量数据处理
```vba
' 分块处理，避免内存溢出
Const CHUNK_SIZE = 5000

For startRow = 2 To totalRows Step CHUNK_SIZE
    endRow = Application.Min(startRow + CHUNK_SIZE - 1, totalRows)
    
    ' 处理当前块
    data = ws.Range("A" & startRow & ":D" & endRow).Value
    ' ... 处理 ...
    ws.Range("A" & startRow & ":D" & endRow).Value = data
    
    Erase data
    DoEvents
Next startRow
```

### 减少Excel交互
```vba
' 批量设置公式后转值
ws.Range("E2:E10000").Formula = "=B2*C2"
ws.Range("E2:E10000").Value = ws.Range("E2:E10000").Value
```

## 安全最佳实践

### 操作前备份
```vba
' 自动创建备份
ActiveSheet.Copy After:=Sheets(Sheets.Count)
ActiveSheet.Name = origSheet.Name & "_备份" & Format(Now, "hhmmss")
```

### 危险操作确认
```vba
If MsgBox("确定要删除吗？", vbYesNo + vbQuestion) = vbYes Then
    ' 执行删除
End If
```

