import os
import re
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
    使用手动匹配括号的方法，提取从 "typedef struct" 到分号结束的整体结构体定义，
    并返回结构体的 typedef 名称和大括号内部的全部内容。
    返回一个列表，每个元素为字典，包含：
      - file: 文件路径
      - typedef_name: 结构体类型名称
      - content: 大括号内的全部内容（字符串）
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
        # 定位 typedef struct
        ts_index = content.find("typedef struct", search_start)
        if ts_index == -1:
            break
        # 定位第一个 '{'
        brace_start = content.find("{", ts_index)
        if brace_start == -1:
            break
        # 利用计数法匹配大括号
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
            search_start = brace_start + 1
            continue
        brace_end = index

        struct_body = content[brace_start+1:brace_end].strip()

        # 在右大括号后查找 typedef 名称，直到遇到分号
        semicolon_index = content.find(";", brace_end)
        if semicolon_index == -1:
            search_start = brace_end + 1
            continue
        typedef_part = content[brace_end+1:semicolon_index].strip()
        if not typedef_part:
            search_start = semicolon_index + 1
            continue
        typedef_name = typedef_part.split()[0]

        structs.append({
            'file': file_path,
            'typedef_name': typedef_name,
            'content': struct_body
        })
        search_start = semicolon_index + 1

    return structs

def parse_struct_members(struct_body):
    """
    将结构体内部内容按分号拆分，解析每一条成员语句。
    对每条成员：
      - 提取块注释（/**/形式）和行注释（//形式）；
      - 去除注释后判断是否为嵌套头文件（#include ...）；
      - 否则尝试匹配变量定义，提取变量类型、变量名称以及数组大小（如存在）。
    返回一个成员列表，每个成员为字典，字段包括：
      - member_code: 原始成员代码（加上分号）
      - var_type: 变量类型（如果能解析到，否则为空）
      - var_name: 变量名称（如果能解析到，否则为空）
      - array_size: 数组大小（如果存在，否则为空）
      - block_comments: 列表，所有块注释内容（/**/）
      - line_comments: 列表，所有行注释内容（//...）
      - include: 嵌套头文件内容（如果该行为 #include 行），否则为空
    """
    members = []
    # 按分号拆分（注意：此处假定分号不会出现在注释内部）
    parts = struct_body.split(';')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        # 恢复分号
        member_line = part + ';'
        # 提取块注释 /**/  (非贪婪)
        block_comments = re.findall(r'/\*.*?\*/', member_line, flags=re.DOTALL)
        # 提取行注释 //...
        line_comments = re.findall(r'//.*', member_line)
        # 去除所有注释，先去除块注释，再去除行注释
        code_no_comments = re.sub(r'/\*.*?\*/', '', member_line, flags=re.DOTALL)
        code_no_comments = re.sub(r'//.*', '', code_no_comments)
        code_no_comments = code_no_comments.strip()
        
        # 判断是否为 #include 行
        include_file = ''
        if code_no_comments.startswith("#include"):
            # 简单提取 #include 后面的内容（支持 "xxx" 或 <xxx>）
            m = re.search(r'#include\s*[<"]([^>"]+)[>"]', code_no_comments)
            if m:
                include_file = m.group(1)
            # 将其他变量信息置空
            var_type = ''
            var_name = ''
            array_size = ''
        else:
            # 尝试解析变量定义
            # 此处采用简单的正则：
            # 例如： "int a", "char b[10]", "unsigned long x", "struct Foo* p"
            # 正则思路：变量类型为连续的字母、数字、下划线、空格和星号（*），后面跟一个变量名，
            # 可选数组部分： [ 数字或其他表达式 ]
            var_type = ''
            var_name = ''
            array_size = ''
            # 匹配模式（注意：此处不支持非常复杂的声明，仅适合简单场景）
            pattern = re.compile(
                r'^(?P<type>[\w\s\*\_]+?)\s+(?P<name>\w+)(\s*\[\s*(?P<size>[^\]]+)\s*\])?\s*$'
            )
            m = pattern.match(code_no_comments)
            if m:
                var_type = m.group('type').strip()
                var_name = m.group('name').strip()
                if m.group('size'):
                    array_size = m.group('size').strip()

        members.append({
            'member_code': member_line,
            'var_type': var_type,
            'var_name': var_name,
            'array_size': array_size,
            'block_comments': block_comments,
            'line_comments': line_comments,
            'include': include_file
        })
    return members

def save_to_excel(all_members, output_excel):
    """
    将所有成员信息保存到 Excel。
    每一行对应一条结构体成员，列包括：
      文件路径, 结构体名称, 成员序号, 成员原始代码, 变量类型, 变量名称, 数组大小, 块注释, 行注释, 嵌套头文件
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "struct_members"
    headers = [
        '文件路径', '结构体名称', '成员序号', '成员原始代码',
        '变量类型', '变量名称', '数组大小',
        '块注释', '行注释', '嵌套头文件'
    ]
    ws.append(headers)

    for member in all_members:
        ws.append([
            member.get('file', ''),
            member.get('struct_name', ''),
            member.get('member_index', ''),
            member.get('member_code', ''),
            member.get('var_type', ''),
            member.get('var_name', ''),
            member.get('array_size', ''),
            "\n".join(member.get('block_comments', [])),
            "\n".join(member.get('line_comments', [])),
            member.get('include', '')
        ])
    try:
        wb.save(output_excel)
        print(f"Excel 文件已生成：{output_excel}")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")

def main():
    # 选择目录对话框
    root = tk.Tk()
    root.withdraw()
    root_dir = filedialog.askdirectory(title="请选择要遍历的根目录")
    if not root_dir:
        print("未选择目录，程序退出。")
        return

    output_excel = 'struct_members.xlsx'
    print(f"开始递归遍历目录 {root_dir} ，查找 .h 文件...")
    h_files = find_h_files_recursive(root_dir)
    print(f"共找到 {len(h_files)} 个 .h 文件。")

    all_members = []
    for file in h_files:
        structs = extract_structs_from_file(file)
        for struct in structs:
            struct_name = struct['typedef_name']
            # 解析结构体成员
            members = parse_struct_members(struct['content'])
            for idx, member in enumerate(members, start=1):
                member['file'] = file
                member['struct_name'] = struct_name
                member['member_index'] = idx
                all_members.append(member)
            print(f"从 {file} 中结构体 {struct_name} 提取 {len(members)} 条成员。")

    if all_members:
        save_to_excel(all_members, output_excel)
    else:
        print("未提取到任何结构体成员。")

if __name__ == "__main__":
    main()