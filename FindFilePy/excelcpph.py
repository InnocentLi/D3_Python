import os
import re
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from concurrent.futures import ThreadPoolExecutor, as_completed

# -------------------------------
# 优化后的正则表达式
# 成员声明：类型由一个或多个标识符（可包含指针符号）组成，变量名必须以字母或下划线开头
# 支持可选数组声明，要求末尾有分号（parse_member 内部保证末尾存在分号）
# -------------------------------
PATTERN_MEMBER = re.compile(
    r'^(?P<type>(?:[a-zA-Z_]\w*\s*(?:\*+\s*)?)+)\s+(?P<name>[a-zA-Z_]\w*)(?:\s*\[\s*(?P<size>[^\]]+)\s*\])?\s*;$'
)
# _firstcall 块：要求名称为有效标识符
PATTERN_FIRSTCALL = re.compile(r'^long\s+_firstcall\s+(?P<name>[a-zA-Z_]\w*)\s*\(')
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
    # 若末尾没有分号，补上（保证正则匹配）
    if not original.endswith(";"):
        original += ";"
    # 提取块注释和行注释
    block_comments = re.findall(r'/\*.*?\*/', original, flags=re.DOTALL)
    line_comments = re.findall(r'//.*', original)
    # 去除注释后的代码部分
    code = re.sub(r'/\*.*?\*/', '', original, flags=re.DOTALL)
    code = re.sub(r'//.*', '', code)
    code = code.strip().rstrip(';').strip()
    # 如果代码仅为 { } ( ) 则跳过
    if code in {"{", "}", "(", ")"}:
        return None
    m = PATTERN_MEMBER.match(code + ";")
    if m:
        var_type = m.group('type').strip()
        var_name = m.group('name').strip()
        array_size = m.group('size').strip() if m.group('size') else ""
    else:
        # 兜底：简单拆分
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
# 逐行处理文件，使用状态机识别块（typedef_struct 或 _firstcall）
# 优化部分：
# 1. 文件编码：先尝试 shift_jis，若失败则改用 utf-8；
# 2. 如果文件结束时块未正常结束，则跳过该块。
# -------------------------------
def process_file(file_path):
    members = []
    current_block_type = None  # "typedef_struct" 或 "fristcall"
    current_block_name = ""
    current_block_lines = []
    
    # 封装一个内部函数用于打开文件（先用 shift_jis，再用 utf-8）
    def open_file_lines(path):
        try:
            return open(path, 'r', encoding='shift_jis')
        except UnicodeDecodeError:
            return open(path, 'r', encoding='utf-8')
    
    try:
        with open_file_lines(file_path) as f:
            for line in f:
                stripped = line.strip()
                # 检测 typedef struct 块开始
                if "typedef struct" in stripped:
                    if current_block_type is not None:
                        # 如果上一个块未正常结束，则跳过它
                        print(f"警告：在 {file_path} 中遇到新块开始，但前一个块未结束，前一块已跳过。")
                        current_block_type = None
                        current_block_lines = []
                    current_block_type = "typedef_struct"
                    current_block_lines = [line]
                    current_block_name = ""
                    continue
                # 检测 _firstcall 块开始
                m_firstcall = PATTERN_FIRSTCALL.search(stripped)
                if m_firstcall:
                    if current_block_type is not None:
                        print(f"警告：在 {file_path} 中遇到新块开始，但前一个块未结束，前一块已跳过。")
                        current_block_type = None
                        current_block_lines = []
                    current_block_type = "fristcall"
                    current_block_name = m_firstcall.group('name').strip()
                    current_block_lines = [line]
                    continue
                # 如果在块内，则累积行
                if current_block_type is not None:
                    current_block_lines.append(line)
                    # 对 typedef_struct 块，遇到右大括号时尝试结束块
                    if current_block_type == "typedef_struct" and "}" in stripped:
                        m = PATTERN_TYPEDEF_STRUCT_END.search(stripped)
                        if m:
                            current_block_name = m.group('name').strip()
                            block_content = "\n".join(current_block_lines)
                            mems = parse_declarations_from_block(block_content)
                            for mem in mems:
                                mem['file'] = file_path
                                mem['block_type'] = "typedef_struct"
                                mem['block_name'] = current_block_name
                            members.extend(mems)
                            current_block_type = None
                            current_block_lines = []
                    # 对 fristcall 块，以遇到以 ");" 结尾的行结束块
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
        # 文件结束后，如果仍处于块内，则说明块未正常结束，直接跳过
        if current_block_type is not None:
            print(f"提示：在 {file_path} 文件末尾，块 {current_block_type} 未正常结束，将跳过该块。")
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