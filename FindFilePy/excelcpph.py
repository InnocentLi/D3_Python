import os
import re
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor, as_completed

# -------------------------------
# 优化后的正则表达式
# 成员声明：要求类型和变量名符合 C 语言标识符规则，允许指针和数组声明，末尾必须有分号
# -------------------------------
PATTERN_MEMBER = re.compile(
    r'^(?P<type>(?:[a-zA-Z_]\w*\s*(?:\*+\s*)?)+)\s+(?P<name>[a-zA-Z_]\w*)'
    r'(?:\s*\[\s*(?P<size>[^\]]+)\s*\])?\s*;$'
)
# 用于提取 _firstcall 块，DOTALL 模式允许跨行匹配
PATTERN_FIRSTCALL_FULL = re.compile(
    r'long\s+_firstcall\s+(?P<name>[a-zA-Z_]\w*)\s*\((?P<content>.*?)\)\s*;',
    re.DOTALL
)
# typedef struct 块结束后提取名称，如 " } MyStruct; "
PATTERN_TYPEDEF_STRUCT_END = re.compile(r'}\s*(?P<name>[a-zA-Z_]\w*)\s*;')

# -------------------------------
# 遍历目录下所有 .h 文件（递归搜索）
# -------------------------------
def find_h_files_recursive(root_dir):
    h_files = []
    for dirpath, _, filenames in os.walk(root_dir):
        for filename in filenames:
            if filename.lower().endswith('.h'):
                h_files.append(os.path.join(dirpath, filename))
    return h_files

# -------------------------------
# 解析单行声明（成员）
# -------------------------------
def parse_member(line):
    original = line.rstrip("\n")
    # 若末尾没有分号，补上便于正则匹配
    if not original.endswith(";"):
        original += ";"
    # 提取块内注释（块注释和行注释）
    block_comments = re.findall(r'/\*.*?\*/', original, flags=re.DOTALL)
    line_comments = re.findall(r'//.*', original)
    # 去除注释后剩余代码
    code = re.sub(r'/\*.*?\*/', '', original, flags=re.DOTALL)
    code = re.sub(r'//.*', '', code)
    code = code.strip().rstrip(';').strip()
    # 如果内容仅为单个括号则跳过
    if code in {"{", "}", "(", ")"}:
        return None
    m = PATTERN_MEMBER.match(code + ";")
    if m:
        var_type = m.group('type').strip()
        var_name = m.group('name').strip()
        array_size = m.group('size').strip() if m.group('size') else ""
    else:
        # 兜底简单拆分
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
        'line_comments': line_comments
    }

# -------------------------------
# 将块内内容按行拆分并解析每一行声明
# -------------------------------
def parse_declarations_from_block(block_content):
    members = []
    for line in block_content.splitlines():
        parsed = parse_member(line)
        if parsed is not None:
            members.append(parsed)
    return members

# -------------------------------
# 读取整个文件内容后进行处理
# 如果块没有正常结束，则跳过该块
# -------------------------------
def process_file(file_path):
    members = []
    # 先尝试 shift_jis 编码，若失败则使用 utf-8
    try:
        try:
            with open(file_path, 'r', encoding='shift_jis', errors='ignore') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
    except Exception as e:
        print(f"读取 {file_path} 时出错: {e}")
        return members

    # 提取 typedef struct 块
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
        # 利用计数法匹配成对大括号
        while index < len(content):
            if content[index] == '{':
                brace_count += 1
            elif content[index] == '}':
                brace_count -= 1
                if brace_count == 0:
                    break
            index += 1
        # 如果未找到匹配结束，则认为该块不完整，直接跳过
        if brace_count != 0:
            print(f"提示：在 {file_path} 中找到 typedef struct 块但未正常结束，跳过。")
            search_start = brace_start + 1
            continue
        brace_end = index
        block_content = content[brace_start + 1: brace_end].strip()
        semicolon_index = content.find(";", brace_end)
        if semicolon_index == -1:
            print(f"提示：在 {file_path} typedef struct 块结束后未找到分号，跳过。")
            search_start = brace_end + 1
            continue
        typedef_part = content[brace_end + 1:semicolon_index].strip()
        if not typedef_part:
            search_start = semicolon_index + 1
            continue
        typedef_name = typedef_part.split()[0]
        mems = parse_declarations_from_block(block_content)
        for mem in mems:
            mem['file'] = file_path
            mem['block_type'] = "typedef_struct"
            mem['block_name'] = typedef_name
            members.append(mem)
        search_start = semicolon_index + 1

    # 提取 _firstcall 块
    for m in PATTERN_FIRSTCALL_FULL.finditer(content):
        block_name = m.group("name").strip()
        block_content = m.group("content").strip()
        mems = parse_declarations_from_block(block_content)
        for mem in mems:
            mem['file'] = file_path
            mem['block_type'] = "fristcall"
            mem['block_name'] = block_name
            members.append(mem)

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
        '块注释', '行注释'
    ]
    ws.append(headers)
    for idx, member in enumerate(all_members, start=1):
        ws.append([
            member.get('file', ''),
            member.get('block_type', ''),
            member.get('block_name', ''),
            idx,
            member.get('member_code', ''),
            member.get('var_type', ''),
            member.get('var_name', ''),
            member.get('array_size', ''),
            "\n".join(member.get('block_comments', [])),
            "\n".join(member.get('line_comments', []))
        ])
    try:
        wb.save(output_excel)
        print(f"Excel 文件已生成：{output_excel}")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")

# -------------------------------
# 主程序入口：采用线程池并行处理文件
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
    with ThreadPoolExecutor(max_workers=4) as executor:
        future_to_file = {executor.submit(process_file, file): file for file in h_files}
        for future in as_completed(future_to_file):
            file = future_to_file[future]
            try:
                file_members = future.result()
                print(f"从 {file} 中提取了 {len(file_members)} 条成员。")
                all_members.extend(file_members)
            except Exception as e:
                print(f"处理 {file} 时发生错误：{e}")
    if all_members:
        save_to_excel(all_members, output_excel)
    else:
        print("未提取到任何块成员。")

if __name__ == "__main__":
    main()