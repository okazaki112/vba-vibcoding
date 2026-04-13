"""
Excel VBA 管理工具 - VBA-vibcoding核心组件
提供 VBA 模块的增删改查、运行等功能

使用方法:
    python vba_manager.py <命令> [参数]

命令:
    list <文件路径>              - 列出所有 VBA 模块
    read <文件路径> <模块名>      - 读取模块代码
    write <文件路径> <模块名> <代码文件>  - 写入/覆盖模块
    delete <文件路径> <模块名>    - 删除模块
    run <文件路径> <宏名> [参数...] - 运行宏
    export <文件路径> <模块名> <输出文件> - 导出模块
    import <文件路径> <模块文件> [新名称] - 导入模块

示例:
    python vba_manager.py list "D:\\test.xlsm"
    python vba_manager.py read "D:\\test.xlsm" "Module1"
    python vba_manager.py write "D:\\test.xlsm" "MyModule" "code.bas"
    python vba_manager.py run "D:\\test.xlsm" "MyModule.Hello"
    python vba_manager.py export "D:\\test.xlsm" "Module1" "backup.bas"

前置要求:
    - pip install pywin32
    - Excel需启用"信任对VBA工程对象模型的访问"
    - 文件必须是.xlsm格式
"""

import win32com.client as win32
import os
import sys
import json


def get_excel_app():
    """获取 Excel 实例（优先使用已打开的）"""
    try:
        excel = win32.GetActiveObject("Excel.Application")
        return excel
    except:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        return excel


def get_workbook(excel, file_path):
    """获取工作簿（优先找已打开的）"""
    file_path = os.path.abspath(file_path)
    
    for wb in excel.Workbooks:
        try:
            if wb.FullName.upper() == file_path.upper() or \
               wb.Name == os.path.basename(file_path):
                return wb
        except:
            continue
    
    # 没找到，打开文件
    if os.path.exists(file_path):
        return excel.Workbooks.Open(file_path)
    else:
        raise FileNotFoundError(f"文件不存在: {file_path}")


def list_modules(file_path):
    """列出所有 VBA 模块"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        vb_project = wb.VBProject
        
        modules = []
        for comp in vb_project.VBComponents:
            type_map = {1: "模块", 2: "类模块", 3: "窗体", 100: "文档"}
            modules.append({
                "name": comp.Name,
                "type": type_map.get(comp.Type, "其他"),
                "type_id": comp.Type,
                "lines": comp.CodeModule.CountOfLines if comp.Type in [1, 2, 100] else 0
            })
        
        return {"success": True, "file": wb.Name, "modules": modules}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def read_module(file_path, module_name):
    """读取模块代码"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        vb_project = wb.VBProject
        
        comp = vb_project.VBComponents(module_name)
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        
        return {"success": True, "module": module_name, "code": code}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def write_module(file_path, module_name, code_file):
    """写入/覆盖模块"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        vb_project = wb.VBProject
        
        # 读取代码文件
        with open(code_file, 'r', encoding='utf-8') as f:
            code = f.read()
        
        # 删除旧模块
        try:
            old_comp = vb_project.VBComponents(module_name)
            vb_project.VBComponents.Remove(old_comp)
        except:
            pass
        
        # 创建新模块
        new_comp = vb_project.VBComponents.Add(1)  # 1=标准模块
        new_comp.Name = module_name
        new_comp.CodeModule.AddFromString(code)
        
        # 保存
        wb.Save()
        
        return {"success": True, "module": module_name, "lines": len(code.splitlines())}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def delete_module(file_path, module_name):
    """删除模块"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        vb_project = wb.VBProject
        
        comp = vb_project.VBComponents(module_name)
        vb_project.VBComponents.Remove(comp)
        wb.Save()
        
        return {"success": True, "deleted": module_name}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def run_macro(file_path, macro_name, *args):
    """运行宏"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        
        # 执行宏
        if args:
            result = excel.Application.Run(macro_name, *args)
        else:
            result = excel.Application.Run(macro_name)
        
        return {"success": True, "macro": macro_name, "result": str(result) if result else None}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def export_module(file_path, module_name, output_file):
    """导出模块"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        vb_project = wb.VBProject
        
        comp = vb_project.VBComponents(module_name)
        comp.Export(output_file)
        
        return {"success": True, "module": module_name, "output": output_file}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def import_module(file_path, module_file, new_name=None):
    """导入模块"""
    excel = get_excel_app()
    
    try:
        wb = get_workbook(excel, file_path)
        vb_project = wb.VBProject
        
        comp = vb_project.VBComponents.Import(module_file)
        if new_name:
            comp.Name = new_name
        
        wb.Save()
        
        return {"success": True, "imported": comp.Name}
    
    except Exception as e:
        return {"success": False, "error": str(e)}


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    command = sys.argv[1]
    
    try:
        if command == "list" and len(sys.argv) >= 3:
            result = list_modules(sys.argv[2])
        
        elif command == "read" and len(sys.argv) >= 4:
            result = read_module(sys.argv[2], sys.argv[3])
        
        elif command == "write" and len(sys.argv) >= 5:
            result = write_module(sys.argv[2], sys.argv[3], sys.argv[4])
        
        elif command == "delete" and len(sys.argv) >= 4:
            result = delete_module(sys.argv[2], sys.argv[3])
        
        elif command == "run" and len(sys.argv) >= 4:
            result = run_macro(sys.argv[2], sys.argv[3], *sys.argv[4:])
        
        elif command == "export" and len(sys.argv) >= 5:
            result = export_module(sys.argv[2], sys.argv[3], sys.argv[4])
        
        elif command == "import" and len(sys.argv) >= 4:
            new_name = sys.argv[4] if len(sys.argv) > 4 else None
            result = import_module(sys.argv[2], sys.argv[3], new_name)
        
        else:
            print(f"❌ 未知命令或参数不足: {command}")
            print(__doc__)
            sys.exit(1)
        
        # 输出 JSON 结果
        print(json.dumps(result, ensure_ascii=False, indent=2))
    
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}, ensure_ascii=False))
        sys.exit(1)


if __name__ == "__main__":
    main()