import os
import re
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

def find_h_files(root_dir):
    """递归遍历目录，返回所有 .h 文件的完整路径列表。"""
    h_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.h'):
                h_files.append(os.path.join(dirpath, filename))
    return h_files

def extract_structs_from_file(file_path):
    """
    从一个 .h 文件中提取所有 typedef struct 定义。
    
    假设结构体的定义格式为：
    
        typedef struct [可选的标签] {
            ... 结构体内容 ...
        } typedef_name;
    
    返回一个列表，每个元素是一个字典，包含：
      - file: 文件路径
      - tag: 可选的标签名称
      - typedef_name: 结构体类型名称
      - content: 大括号内的内容
    """
    structs = []
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"读取 {file_path} 时出错: {e}")
        return structs

    pattern = re.compile(
        r'typedef\s+struct\s*(\w+)?\s*\{(.*?)\}\s*(\w+)\s*;',
        re.DOTALL
    )

    for match in pattern.finditer(content):
        struct_info = {
            'file': file_path,
            'tag': match.group(1) if match.group(1) else '',
            'typedef_name': match.group(3),
            'content': match.group(2).strip()
        }
        structs.append(struct_info)
    return structs

def save_to_excel(structs, output_excel):
    """将结构体信息保存到 Excel 文件中。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "typedef_structs"

    # 设置表头
    headers = ['文件路径', '标签', 'typedef名称', '结构体内容']
    ws.append(headers)

    for struct in structs:
        ws.append([
            struct['file'],
            struct['tag'],
            struct['typedef_name'],
            struct['content']
        ])

    try:
        wb.save(output_excel)
        print(f"Excel 文件已生成：{output_excel}")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")

def main():
    # 使用 tkinter.filedialog 弹出选择目录对话框
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root_dir = filedialog.askdirectory(title="请选择要遍历的根目录")
    if not root_dir:
        print("未选择目录，程序退出。")
        return

    output_excel = 'typedef_structs.xlsx'

    print(f"开始遍历目录 {root_dir} ，查找 .h 文件...")
    h_files = find_h_files(root_dir)
    print(f"共找到 {len(h_files)} 个 .h 文件。")

    all_structs = []
    for file in h_files:
        structs = extract_structs_from_file(file)
        if structs:
            print(f"从 {file} 中提取到 {len(structs)} 个 typedef struct 定义。")
            all_structs.extend(structs)

    if all_structs:
        save_to_excel(all_structs, output_excel)
    else:
        print("未提取到任何 typedef struct 定义。")

if __name__ == "__main__":
    main()