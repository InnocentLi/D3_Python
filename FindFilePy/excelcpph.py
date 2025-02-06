import os
import re
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

# 解析单行声明，保留注释、提取变量类型、变量名称、数组大小
def parse_member(line):
    # 原始声明（确保末尾有分号，方便观察）
    original = line.rstrip("\n")
    if not original.endswith(";"):
        original += ";"
    # 提取块注释 /**/ 和行注释 //  
    block_comments = re.findall(r'/\*.*?\*/', original, flags=re.DOTALL)
    line_comments = re.findall(r'//.*', original)
    # 去除注释后的代码部分
    code = re.sub(r'/\*.*?\*/', '', original, flags=re.DOTALL)
    code = re.sub(r'//.*', '', code)
    code = code.strip().rstrip(';').strip()
    # 如果代码以 "#include" 开头，则视为嵌套头文件引用
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
    # 尝试解析变量声明
    # 采用较宽松的正则，不限定必须以某些关键字开头
    pattern = re.compile(r'^(?P<type>[\w\*\s]+?)\s+(?P<name>\*?\w+)(\s*\[\s*(?P<size>[^\]]+)\s*\])?$')
    m = pattern.match(code)
    if m:
        var_type = m.group('type').strip()
        var_name = m.group('name').strip()
        array_size = m.group('size').strip() if m.group('size') else ""
    else:
        # 简单分词：认为第一个 token为类型，其余为变量名
        tokens = code.split()
        if len(tokens) >= 2:
            var_type = tokens[0]
            var_name = tokens[1]
            array_size = ""
        else:
            var_type, var_name, array_size = code, "", ""
    return {
        'member_code': original,
        'var_type': var_type,
        'var_name': var_name,
        'array_size': array_size,
        'block_comments': block_comments,
        'line_comments': line_comments,
        'include': ""
    }

# 处理一个块，按真实行解析块内的成员声明
def process_block(block_type, block_name, block_lines, file_path):
    members = []
    # 对于 typedef_struct 块，假设块内容位于大括号内
    # 对于 long _firstcall 块，第一行为块头，后续非空行为成员
    content_lines = []
    if block_type == "typedef_struct":
        inside = False
        for line in block_lines:
            if not inside and "{" in line:
                # 从 "{" 后开始
                pos = line.find("{")
                tail = line[pos+1:]
                if tail.strip():
                    content_lines.append(tail)
                inside = True
            elif inside:
                # 如果行中含有 "}"，则只取 "}" 前面的部分
                if "}" in line:
                    pos = line.find("}")
                    head = line[:pos]
                    if head.strip():
                        content_lines.append(head)
                    # 块结束，忽略余下内容
                    break
                else:
                    content_lines.append(line)
    elif block_type == "fristcall":
        # 第一行为块头，后续非空行为成员
        if len(block_lines) >= 2:
            content_lines = block_lines[1:]
        else:
            content_lines = []
    # 按行处理，每行视为一条声明（非空行）
    member_index = 1
    for line in content_lines:
        if line.strip() == "":
            continue
        parsed = parse_member(line)
        # 补充块及文件信息
        parsed['file'] = file_path
        parsed['block_type'] = block_type
        parsed['block_name'] = block_name
        parsed['member_index'] = member_index
        member_index += 1
        members.append(parsed)
    return members

# 按行扫描文件，使用状态机识别块并解析块内成员
def process_file(file_path):
    members = []
    current_block_type = None    # "typedef_struct" 或 "fristcall"
    current_block_name = ""
    current_block_lines = []
    with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
        lines = f.readlines()
    # 按行扫描
    for line in lines:
        stripped = line.strip()
        # 判断是否为 typedef struct 块头
        if "typedef struct" in stripped:
            # 如果已有块未结束，则先处理前一个块
            if current_block_type is not None:
                members.extend(process_block(current_block_type, current_block_name, current_block_lines, file_path))
                current_block_type = None
                current_block_lines = []
            current_block_type = "typedef_struct"
            current_block_lines = [line]
            current_block_name = ""
            continue
        # 判断是否为 long _firstcall 块头（注意使用英文圆括号）
        m_firstcall = re.search(r'^long\s+_firstcall\s*\(\s*([^\)]+)\s*\)', stripped)
        if m_firstcall:
            if current_block_type is not None:
                members.extend(process_block(current_block_type, current_block_name, current_block_lines, file_path))
                current_block_type = None
                current_block_lines = []
            current_block_type = "fristcall"
            current_block_lines = [line]
            current_block_name = m_firstcall.group(1).strip()
            continue
        # 如果处于块内，则累计行
        if current_block_type is not None:
            current_block_lines.append(line)
            # 对 typedef_struct 块，当行中出现 "}" 并且含有 typedef 名称时认为结束
            if current_block_type == "typedef_struct" and "}" in line:
                # 尝试从行中提取 typedef 名称，例如 "} Name;"
                m = re.search(r'}\s*(\w+)\s*;', line)
                if m:
                    current_block_name = m.group(1).strip()
                # 认为块结束，处理该块
                members.extend(process_block(current_block_type, current_block_name, current_block_lines, file_path))
                current_block_type = None
                current_block_lines = []
            # 对 fristcall 块，遇到空行则认为块结束
            elif current_block_type == "fristcall" and stripped == "":
                members.extend(process_block(current_block_type, current_block_name, current_block_lines, file_path))
                current_block_type = None
                current_block_lines = []
        # 若不在块中，忽略当前行
    # 文件结束时，如有未处理的块，则处理之
    if current_block_type is not None and current_block_lines:
        members.extend(process_block(current_block_type, current_block_name, current_block_lines, file_path))
    return members

def find_h_files_recursive(root_dir):
    h_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.h'):
                h_files.append(os.path.join(dirpath, filename))
    return h_files

def save_to_excel(all_members, output_excel):
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
        file_members = process_file(file)
        print(f"从 {file} 中提取了 {len(file_members)} 条成员。")
        all_members.extend(file_members)
    if all_members:
        save_to_excel(all_members, output_excel)
    else:
        print("未提取到任何块成员。")

if __name__ == "__main__":
    main()