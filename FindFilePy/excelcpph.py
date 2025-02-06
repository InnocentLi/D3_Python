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
    采用手动匹配大括号的方法，从 "typedef struct" 开始，
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

def parse_struct_members_by_basictype(struct_body):
    """
    按照 C 语言的基本类型进行解析：
      1. 利用正则表达式，寻找以 int、char、float、double、short、long、
         unsigned、signed、struct、union、enum 开头的声明串，直到遇到分号；
      2. 对于每个匹配到的声明串：
         - 提取其中所有的块注释（/**/）和行注释（//）；
         - 去除注释后，分离出“基础类型”和变量部分；
         - 变量部分可能包含多个变量，以逗号分隔，每个变量可能有指针符号和数组定义；
         - 分别解析出变量名称和数组大小（如果存在），并将指针符号附加到基础类型中；
      3. 返回解析得到的所有成员列表，每个成员为一个字典，包含原始声明、变量类型、变量名称、数组大小、块注释、行注释。
    """
    members = []
    # 匹配以基本类型开头的声明（注意：\b 确保完整单词匹配；.*? 非贪婪匹配到第一个分号）
    pattern = re.compile(
        r'(?P<decl>\s*(?:(?:int|char|float|double|short|long|unsigned|signed|struct|union|enum)\b.*?;))',
        re.DOTALL | re.MULTILINE
    )
    for match in pattern.finditer(struct_body):
        decl = match.group('decl')
        # 提取注释
        block_comments = re.findall(r'/\*.*?\*/', decl, flags=re.DOTALL)
        line_comments = re.findall(r'//.*', decl)
        # 去除所有注释，方便后续解析
        decl_no_comments = re.sub(r'/\*.*?\*/', '', decl, flags=re.DOTALL)
        decl_no_comments = re.sub(r'//.*', '', decl_no_comments)
        decl_no_comments = decl_no_comments.strip().rstrip(';').strip()
        # 利用正则分离基础类型与变量部分
        m = re.match(
            r'^(?P<base>(?:int|char|float|double|short|long|unsigned|signed|struct|union|enum)(?:[\w\s\*\_]*))\s+(?P<vars>.+)$',
            decl_no_comments
        )
        if not m:
            continue
        base_type = m.group('base').strip()
        vars_str = m.group('vars').strip()
        # 将变量部分以逗号分割（针对类似 int a, *b, c[10] 的情况）
        var_tokens = [t.strip() for t in vars_str.split(',')]
        for token in var_tokens:
            array_size = ''
            # 如果变量中包含数组定义，如 a[10]
            m_array = re.search(r'\[(.*?)\]', token)
            if m_array:
                array_size = m_array.group(1).strip()
                token = re.sub(r'\[.*?\]', '', token).strip()
            # 提取变量名称，若有指针符号则剥离
            pointer_prefix = ''
            while token.startswith('*'):
                pointer_prefix += '*'
                token = token[1:].strip()
            var_name = token
            # 如果存在指针符号，则类型加上
            var_type = base_type if not pointer_prefix else f"{base_type} {pointer_prefix}"
            members.append({
                'member_code': decl.strip(),  # 原始声明字符串
                'var_type': var_type,
                'var_name': var_name,
                'array_size': array_size,
                'block_comments': block_comments,
                'line_comments': line_comments,
                'include': ''  # 本函数暂不处理 #include 嵌套
            })
    return members

def save_to_excel(all_members, output_excel):
    """
    将所有结构体成员信息保存到 Excel 文件中。
    每一行对应一条结构体成员，列包括：
      文件路径、结构体名称、成员序号、原始声明、变量类型、变量名称、数组大小、
      块注释、行注释、嵌套头文件（若有）
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "struct_members"
    headers = [
        '文件路径', '结构体名称', '成员序号', '原始声明',
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
    # 弹出目录选择对话框
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
            # 使用按基本类型解析成员的方法
            members = parse_struct_members_by_basictype(struct['content'])
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