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

def extract_typedef_structs_from_file(file_path):
    """
    从文件中提取 typedef struct 块：
      模式：typedef struct [optional_tag] { ... } typedef_name;
    返回列表，每个元素为字典，包含：
      - file: 文件路径
      - block_type: 固定为 "typedef_struct"
      - name: typedef名称
      - content: 大括号内内容
    """
    blocks = []
    try:
        with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"读取 {file_path} 时出错: {e}")
        return blocks

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
        block_content = content[brace_start+1:brace_end].strip()
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
        blocks.append({
            'file': file_path,
            'block_type': 'typedef_struct',
            'name': typedef_name,
            'content': block_content
        })
        search_start = semicolon_index + 1
    return blocks

def extract_fristcalls_from_file(file_path):
    """
    从文件中提取 long _fristcall 块：
      格式：long _fristcall( someName ) { ... }
    返回列表，每个元素为字典，包含：
      - file: 文件路径
      - block_type: 固定为 "fristcall"
      - name: 括号内提取的名称
      - content: 大括号内内容
    """
    blocks = []
    try:
        with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"读取 {file_path} 时出错: {e}")
        return blocks

    # 匹配 long _fristcall( name ) 后跟左大括号
    pattern = re.compile(r'long\s+_fristcall\s*\(\s*(?P<name>[^\)]+)\s*\)\s*\{', re.MULTILINE)
    search_start = 0
    while True:
        m = pattern.search(content, search_start)
        if not m:
            break
        call_name = m.group('name').strip()
        brace_start = content.find("{", m.start())
        if brace_start == -1:
            search_start = m.end()
            continue
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
        block_content = content[brace_start+1:brace_end].strip()
        blocks.append({
            'file': file_path,
            'block_type': 'fristcall',
            'name': call_name,
            'content': block_content
        })
        search_start = brace_end + 1
    return blocks

def parse_declaration(code):
    """
    尝试解析单个变量声明（不含末尾分号）的字符串，
    返回 (var_type, var_name, array_size)。
    使用较宽松的正则，不再限定必须以 int/char 开头。
    如果匹配失败，做简单拆分：认为第一个 token 为类型，其余为变量名。
    """
    pattern = re.compile(r'^(?P<type>[\w\*\s]+?)\s+(?P<name>\*?\w+)(\s*\[\s*(?P<size>[^\]]+)\s*\])?$')
    m = pattern.match(code)
    if m:
        var_type = m.group('type').strip()
        var_name = m.group('name').strip()
        array_size = m.group('size').strip() if m.group('size') else ""
        return var_type, var_name, array_size
    else:
        tokens = code.split()
        if len(tokens) >= 2:
            return tokens[0], tokens[1], ""
        else:
            return code, "", ""

def parse_member(decl_str):
    """
    针对单个声明字符串（不含末尾分号），解析出：
      - 原始声明（加上分号）
      - 变量类型、变量名称、数组大小
      - 提取块注释和行注释
      - 如果去除注释后以 "#include" 开头，则视为嵌套头文件引用
    """
    original = decl_str.strip() + ";"
    block_comments = re.findall(r'/\*.*?\*/', original, flags=re.DOTALL)
    line_comments = re.findall(r'//.*', original)
    # 去除注释后用于解析代码部分
    code = re.sub(r'/\*.*?\*/', '', original, flags=re.DOTALL)
    code = re.sub(r'//.*', '', code)
    code = code.strip().rstrip(';').strip()
    # 判断是否为嵌套的 #include 行
    if code.startswith("#include"):
        m = re.search(r'#include\s*[<"]([^>"]+)[>"]', code)
        include = m.group(1) if m else ""
        return {
            'member_code': original,
            'var_type': "",
            'var_name': "",
            'array_size': "",
            'block_comments': block_comments,
            'line_comments': line_comments,
            'include': include
        }
    var_type, var_name, array_size = parse_declaration(code)
    return {
        'member_code': original,
        'var_type': var_type,
        'var_name': var_name,
        'array_size': array_size,
        'block_comments': block_comments,
        'line_comments': line_comments,
        'include': ""
    }

def parse_declarations_from_block(block_content):
    """
    针对块内代码进行成员解析：
      1. 先按分号拆分（假设分号不在注释中）
      2. 对每个部分，若其中存在逗号，则认为是一条多变量声明，
         则采用“类型”部分+逗号分隔的各个变量进行逐个解析；
      3. 如果不含逗号，则直接解析为一条声明。
    返回成员列表，每个成员为 parse_member() 的结果字典。
    """
    members = []
    # 按分号拆分
    parts = block_content.split(';')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        # 若包含逗号，则认为前面部分为类型，后面为多个变量
        if ',' in part:
            # 用正则分离类型和变量部分
            m = re.match(r'^(?P<type>[\w\*\s]+?)\s+(?P<vars>.+)$', part)
            if m:
                base_type = m.group('type').strip()
                vars_part = m.group('vars').strip()
                var_tokens = [v.strip() for v in vars_part.split(',')]
                for token in var_tokens:
                    full_decl = base_type + " " + token
                    member = parse_member(full_decl)
                    members.append(member)
            else:
                # 若匹配失败，则整体解析
                member = parse_member(part)
                members.append(member)
        else:
            member = parse_member(part)
            members.append(member)
    return members

def save_to_excel(all_members, output_excel):
    """
    将所有成员信息保存到 Excel 文件中。
    每一行对应一条成员，字段包括：
      文件路径、块类型、块名称、成员序号、原始声明、变量类型、变量名称、
      数组大小、块注释、行注释、嵌套头文件
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "struct_members"
    headers = [
        '文件路径', '块类型', '块名称', '成员序号', '原始声明',
        '变量类型', '变量名称', '数组大小',
        '块注释', '行注释', '嵌套头文件'
    ]
    ws.append(headers)
    for member in all_members:
        ws.append([
            member.get('file', ''),
            member.get('block_type', ''),
            member.get('block_name', ''),
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
        # 提取 typedef struct 块
        typedef_blocks = extract_typedef_structs_from_file(file)
        # 提取 long _fristcall 块
        fristcall_blocks = extract_fristcalls_from_file(file)
        # 合并两类块
        blocks = []
        for b in typedef_blocks:
            # 对 typedef_struct 块，块名称取 typedef_name
            b['block_type'] = 'typedef_struct'
            b['block_name'] = b['name']
            blocks.append(b)
        for b in fristcall_blocks:
            b['block_type'] = 'fristcall'
            b['block_name'] = b['name']
            blocks.append(b)
        # 针对每个块，解析成员
        for blk in blocks:
            members = parse_declarations_from_block(blk['content'])
            for idx, member in enumerate(members, start=1):
                member['file'] = file
                member['block_type'] = blk['block_type']
                member['block_name'] = blk['block_name']
                member['member_index'] = idx
                all_members.append(member)
            print(f"从 {file} 中块 {blk['block_type']}({blk['block_name']}) 提取 {len(members)} 条成员。")

    if all_members:
        save_to_excel(all_members, output_excel)
    else:
        print("未提取到任何结构体/块成员。")

if __name__ == "__main__":
    main()