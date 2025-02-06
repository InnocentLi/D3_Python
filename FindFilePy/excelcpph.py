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
    采用手动匹配大括号的方法，从 "typedef struct" 开始匹配，
    提取整个结构体内部的代码和 typedef 名称。
    返回的每个结果为字典，包含文件路径、结构体名称和大括号内的全部内容。
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
        ts_index = content.find("typedef struct", search_start)
        if ts_index == -1:
            break
        brace_start = content.find("{", ts_index)
        if brace_start == -1:
            break
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

def parse_variable_declaration(code):
    """
    针对去除注释后的单行变量声明进行简单解析：
      - 去除末尾的分号后，以空白拆分。
      - 认为最后一个 token 为变量名（可能含有数组声明，如 a[10]），
        其余部分合并为变量类型。
      - 如果变量名中含有 '['，则提取数组大小。
    返回 (var_type, var_name, array_size) 三个字符串。
    """
    # 去除末尾分号与多余空白
    code = code.strip().rstrip(';').strip()
    if not code:
        return '', '', ''
    tokens = code.split()
    if not tokens:
        return '', '', ''
    last_token = tokens[-1]
    array_size = ''
    # 如果最后一个 token 中包含数组声明，如 a[10]
    if '[' in last_token:
        var_name = last_token.split('[')[0]
        m = re.search(r'\[\s*([^\]]+)\s*\]', last_token)
        if m:
            array_size = m.group(1)
    else:
        var_name = last_token
    var_type = ' '.join(tokens[:-1])
    return var_type, var_name, array_size

def parse_struct_members(struct_body):
    """
    将结构体内部内容按分号拆分，逐条解析每个成员。
    对每个成员：
      1. 提取块注释（/**/形式）和行注释（//形式）。
      2. 去除所有注释后检查是否为嵌套的 #include 行。
      3. 否则采用 parse_variable_declaration() 解析变量类型、变量名称和数组大小。
    返回成员列表，每个成员为字典，包含：
      - member_code: 原始成员代码（带分号）
      - var_type: 变量类型（解析到的）
      - var_name: 变量名称（解析到的）
      - array_size: 数组大小（如存在）
      - block_comments: 列表，所有块注释内容
      - line_comments: 列表，所有行注释内容
      - include: 嵌套头文件（如果该行为 #include 行，否则为空）
    """
    members = []
    # 按分号拆分；注意：此处假定分号不会出现在注释内部
    parts = struct_body.split(';')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        member_line = part + ';'
        # 提取块注释 /**/（非贪婪模式）
        block_comments = re.findall(r'/\*.*?\*/', member_line, flags=re.DOTALL)
        # 提取行注释 //...
        line_comments = re.findall(r'//.*', member_line)
        # 去除注释：先去掉块注释，再去掉行注释
        code_no_comments = re.sub(r'/\*.*?\*/', '', member_line, flags=re.DOTALL)
        code_no_comments = re.sub(r'//.*', '', code_no_comments)
        code_no_comments = code_no_comments.strip()
        
        include_file = ''
        var_type = ''
        var_name = ''
        array_size = ''
        # 如果为嵌套头文件声明
        if code_no_comments.startswith("#include"):
            m = re.search(r'#include\s*[<"]([^>"]+)[>"]', code_no_comments)
            if m:
                include_file = m.group(1)
        else:
            # 如果代码中不为空，则尝试解析变量声明
            if code_no_comments:
                var_type, var_name, array_size = parse_variable_declaration(code_no_comments)
        
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
    将所有结构体成员信息保存到 Excel 文件中。
    每一行对应一条结构体成员，列包括：
      文件路径、结构体名称、成员序号、原始成员代码、变量类型、变量名称、数组大小、
      块注释、行注释、嵌套头文件
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
    # 弹出选择目录对话框
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