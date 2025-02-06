import os
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

def find_h_files_recursive(root_dir):
    """
    递归遍历目录，返回所有 .h 文件的完整路径列表。
    """
    h_files = []
    try:
        with os.scandir(root_dir) as it:
            for entry in it:
                if entry.is_file() and entry.name.lower().endswith('.h'):
                    h_files.append(entry.path)
                elif entry.is_dir():
                    h_files.extend(find_h_files_recursive(entry.path))
    except Exception as e:
        print(f"遍历 {root_dir} 时出错: {e}")
    return h_files

def extract_structs_from_file(file_path):
    """
    从一个 .h 文件中提取所有 typedef struct 定义。
    
    使用手动解析方法，从 "typedef struct" 开始，
    找到第一个 '{'，然后利用括号计数法获取匹配的 '}' 所在位置，
    接着提取大括号内的全部内容，再提取 typedef 名称（直到遇到分号）。
    
    返回一个列表，每个元素是一个字典，包含：
      - file: 文件路径
      - typedef_name: 结构体类型名称
      - content: 大括号内的全部内容
    """
    structs = []
    try:
        with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"读取 {file_path} 时出错: {e}")
        return structs

    search_start = 0
    while True:
        # 找到 "typedef struct" 出现的位置
        ts_index = content.find("typedef struct", search_start)
        if ts_index == -1:
            break  # 没有更多的 typedef struct
        # 从 typedef struct 后查找第一个 '{'
        brace_start = content.find("{", ts_index)
        if brace_start == -1:
            break  # 无法找到左大括号
        # 开始括号匹配
        index = brace_start
        brace_count = 0
        while index < len(content):
            if content[index] == '{':
                brace_count += 1
            elif content[index] == '}':
                brace_count -= 1
                if brace_count == 0:
                    break
            index += 1
        if brace_count != 0:
            # 未能正确匹配到对应的 '}'，跳过本次搜索
            search_start = brace_start + 1
            continue

        brace_end = index  # 对应 '}' 的位置

        # 提取大括号内的全部内容
        struct_body = content[brace_start + 1:brace_end].strip()

        # 在右大括号后查找 typedef 名称，直到遇到分号 ';'
        semicolon_index = content.find(";", brace_end)
        if semicolon_index == -1:
            search_start = brace_end + 1
            continue
        # typedef 名称一般位于 '}' 与 ';' 之间
        typedef_part = content[brace_end + 1:semicolon_index].strip()
        # typedef 名称可能还夹带一些空白或换行，取第一个非空部分作为名称
        if not typedef_part:
            search_start = semicolon_index + 1
            continue
        typedef_name = typedef_part.split()[0]

        struct_info = {
            'file': file_path,
            'typedef_name': typedef_name,
            'content': struct_body
        }
        structs.append(struct_info)

        # 继续查找下一个 typedef struct，从当前分号后开始
        search_start = semicolon_index + 1

    return structs

def save_to_excel(structs, output_excel):
    """将提取的结构体信息保存到 Excel 文件中。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "typedef_structs"

    # 写入表头
    headers = ['文件路径', 'typedef名称', '结构体内容']
    ws.append(headers)

    for struct in structs:
        ws.append([
            struct['file'],
            struct['typedef_name'],
            struct['content']
        ])

    try:
        wb.save(output_excel)
        print(f"Excel 文件已生成：{output_excel}")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")

def main():
    # 弹出选择目录对话框
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root_dir = filedialog.askdirectory(title="请选择要遍历的根目录")
    if not root_dir:
        print("未选择目录，程序退出。")
        return

    output_excel = 'typedef_structs.xlsx'

    print(f"开始递归遍历目录 {root_dir} ，查找 .h 文件...")
    h_files = find_h_files_recursive(root_dir)
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