import os
import re
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
import concurrent.futures

# -------------------------------
# 预编译常用正则表达式
# -------------------------------
PATTERN_INCLUDE = re.compile(r'#include\s*[<"]([^>"]+)[>"]')
PATTERN_MEMBER = re.compile(
    r'^(?P<type>(?:\w+\s*\*?\s*)+)\s+(?P<name>\w+)(?:\s*\[\s*(?P<size>[^\]]+)\s*\])?$'
)
PATTERN_FIRSTCALL = re.compile(r'^long\s+_firstcall\s+(\w+)\s*\(')
PATTERN_TYPEDEF_STRUCT_END = re.compile(r'}\s*(\w+)\s*;')

# -------------------------------
# 文件遍历函数
# -------------------------------
def find_h_files_recursive(root_dir):
    h_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.h'):
                h_files.append(os.path.join(dirpath, filename))
    return h_files

# -------------------------------
# 提取 typedef struct 块
# -------------------------------
def extract_typedef_structs_from_file(file_path):
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
        # 提取 typedef 名称：在右大括号后查找直到分号处的部分
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

# -------------------------------
# 提取 long _firstcall 块
# -------------------------------
def extract_fristcalls_from_file(file_path):
    blocks = []
    try:
        with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"读取 {file_path} 时出错: {e}")
        return blocks

    search_start = 0
    # 匹配形式：long _firstcall MyBlockName( ... );
    pattern = re.compile(r'long\s+_firstcall\s+(\w+)\s*\(', re.MULTILINE)
    while True:
        m = pattern.search(content, search_start)
        if not m:
            break
        block_name = m.group(1).strip()
        paren_start = content.find("(", m.end()-1)
        if paren_start == -1:
            search_start = m.end()
            continue
        index = paren_start
        paren_count = 0
        while index < len(content):
            if content[index] == '(':
                paren_count += 1
            elif content[index] == ')':
                paren_count -= 1
                if paren_count == 0:
                    break
            index += 1
        if paren_count != 0:
            search_start = paren_start + 1
            continue
        paren_end = index
        block_content = content[paren_start+1:paren_end].strip()
        # 判断是否有分号结尾
        semicolon_index = content.find(";", paren_end)
        if semicolon_index == -1:
            search_start = paren_end + 1
            continue

        blocks.append({
            'file': file_path,
            'block_type': 'fristcall',
            'name': block_name,
            'content': block_content
        })
        search_start = semicolon_index + 1
    return blocks

# -------------------------------
# 解析单行声明（成员）
# -------------------------------
def parse_member(line):
    original = line.rstrip("\n")
    if not original.endswith(";"):
        original += ";"
    # 预先提取注释
    block_comments = re.findall(r'/\*.*?\*/', original, flags=re.DOTALL)
    line_comments = re.findall(r'//.*', original)
    # 去除注释后的代码部分
    code = re.sub(r'/\*.*?\*/', '', original, flags=re.DOTALL)
    code = re.sub(r'//.*', '', code)
    code = code.strip().rstrip(';').strip()
    # 如果以 "#include" 开头，视为嵌套头文件引用
    if code.startswith("#include"):
        m = PATTERN_INCLUDE.search(code)
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
    # 使用较宽松的正则表达式捕获类型和变量名
    m = PATTERN_MEMBER.match(code)
    if m:
        var_type = m.group('type').strip()
        var_name = m.group('name').strip()
        array_size = m.group('size').strip() if m.group('size') else ""
    else:
        tokens = code.split()
        if len(tokens) >= 2:
            var_type = " ".join(tokens[:-1])
            var_name = tokens[-1]
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

# -------------------------------
# 将块内内容按行拆分并解析每一行声明
# -------------------------------
def parse_declarations_from_block(block_content):
    # 利用生成器过滤掉空行
    lines = (line for line in block_content.splitlines() if line.strip() != "")
    members = []
    member_index = 1
    for line in lines:
        parsed = parse_member(line)
        parsed['member_index'] = member_index
        member_index += 1
        members.append(parsed)
    return members

# -------------------------------
# 并行处理时的文件包装函数
# -------------------------------
def process_file(file_path):
    """
    逐行处理文件，通过状态机识别块（typedef_struct 或 fristcall），
    对块内内容调用 parse_declarations_from_block() 解析成员。
    此处采用按行扫描方式，防止一次性读取整个文件导致字符串拼接耗时。
    """
    members = []
    current_block_type = None  # "typedef_struct" 或 "fristcall"
    current_block_name = ""
    current_block_lines = []
    try:
        with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
            for line in f:
                stripped = line.strip()
                # 判断是否为 typedef struct 块头
                if "typedef struct" in stripped:
                    if current_block_type is not None:
                        block_content = "\n".join(current_block_lines)
                        mems = parse_declarations_from_block(block_content)
                        for mem in mems:
                            mem['file'] = file_path
                            mem['block_type'] = current_block_type
                            mem['block_name'] = current_block_name
                        members.extend(mems)
                        current_block_type = None
                        current_block_lines = []
                    current_block_type = "typedef_struct"
                    current_block_lines = [line]
                    current_block_name = ""
                    continue
                # 判断是否为 long _firstcall 块头
                m_firstcall = PATTERN_FIRSTCALL.search(stripped)
                if m_firstcall:
                    if current_block_type is not None:
                        block_content = "\n".join(current_block_lines)
                        mems = parse_declarations_from_block(block_content)
                        for mem in mems:
                            mem['file'] = file_path
                            mem['block_type'] = current_block_type
                            mem['block_name'] = current_block_name
                        members.extend(mems)
                        current_block_type = None
                        current_block_lines = []
                    current_block_type = "fristcall"
                    current_block_name = m_firstcall.group(1).strip()
                    current_block_lines = [line]
                    continue
                # 如果在块内，则累计行
                if current_block_type is not None:
                    current_block_lines.append(line)
                    if current_block_type == "typedef_struct" and "}" in stripped:
                        m = PATTERN_TYPEDEF_STRUCT_END.search(stripped)
                        if m:
                            current_block_name = m.group(1).strip()
                        block_content = "\n".join(current_block_lines)
                        mems = parse_declarations_from_block(block_content)
                        for mem in mems:
                            mem['file'] = file_path
                            mem['block_type'] = "typedef_struct"
                            mem['block_name'] = current_block_name
                        members.extend(mems)
                        current_block_type = None
                        current_block_lines = []
                    elif current_block_type == "fristcall" and stripped.endswith(");"):
                        block_content = "\n".join(current_block_lines)
                        mems = parse_declarations_from_block(block_content)
                        for mem in mems:
                            mem['file'] = file_path
                            mem['block_type'] = "fristcall"
                            mem['block_name'] = current_block_name
                        members.extend(mems)
                        current_block_type = None
                        current_block_lines = []
        # 文件末尾处理未结束块
        if current_block_type is not None and current_block_lines:
            block_content = "\n".join(current_block_lines)
            mems = parse_declarations_from_block(block_content)
            for mem in mems:
                mem['file'] = file_path
                mem['block_type'] = current_block_type
                mem['block_name'] = current_block_name
            members.extend(mems)
    except Exception as e:
        print(f"处理文件 {file_path} 时出错：{e}")
    return members

# -------------------------------
# 将所有成员信息写入 Excel
# -------------------------------
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

# -------------------------------
# 主程序入口（并行处理所有 .h 文件）
# -------------------------------
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
    # 使用线程池并行处理文件
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(process_file, file): file for file in h_files}
        for future in concurrent.futures.as_completed(futures):
            file = futures[future]
            try:
                members = future.result()
                print(f"从 {file} 中提取了 {len(members)} 条成员。")
                all_members.extend(members)
            except Exception as e:
                print(f"处理文件 {file} 时出错：{e}")
    if all_members:
        save_to_excel(all_members, output_excel)
    else:
        print("未提取到任何块成员。")

if __name__ == "__main__":
    main()