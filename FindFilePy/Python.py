import os
import re
import openpyxl
from openpyxl import Workbook

import tkinter as tk
from tkinter import filedialog

def parse_cma_file(file_path):
    """
    解析单个 .cma 文件，返回所有 *LINE 块。
    每个块会以(块内容字符串)的形式返回，方便后续再做键值对提取。
    """
    results = []
    # 使用Shift_JIS编码来读取文件
    with open(file_path, 'r', encoding='shift_jis', errors='replace') as f:
        lines = f.readlines()
    
    # 去掉首尾空格和换行
    lines = [line.strip() for line in lines]
    
    current_block = []
    collecting_block = False

    for line in lines:
        # 如果是 *LINE 开头，说明新块开始
        if line.startswith("*LINE"):
            # 如果上一个块还未结束，则先把它存起来
            if collecting_block and current_block:
                combined_line = " ".join(current_block)
                results.append(combined_line)
                current_block = []
            
            collecting_block = True
            current_block.append(line)
        else:
            # 如果正在收集块，则将本行加入
            if collecting_block:
                current_block.append(line)
                # 如果遇到分号, 表示块结束
                if ";" in line:
                    combined_line = " ".join(current_block)
                    results.append(combined_line)
                    collecting_block = False
                    current_block = []
    
    # 如果文件末尾还有没结束的块，也存一下
    if collecting_block and current_block:
        combined_line = " ".join(current_block)
        results.append(combined_line)
    
    return results

def extract_key_values(block_str):
    """
    给定一个块的字符串，比如 "*LINE A=xxx B=yyy ;"
    提取其中的所有 "key=value" 对，返回一个 dict。
    例如返回 {"A": "xxx", "B": "yyy"}
    """
    # 1) 去掉 *LINE
    block_str = block_str.replace("*LINE", "")
    # 2) 按分号拆分，只保留分号前面的部分
    block_str = block_str.split(";")[0]
    block_str = block_str.strip()
    
    # 用正则找出所有 k=v 形式
    pattern = r'(\w+)\s*=\s*([^=\s]+)'
    pairs = re.findall(pattern, block_str)
    
    # 生成字典
    kv_dict = {}
    for k, v in pairs:
        kv_dict[k] = v
    return kv_dict

def scan_cma_and_collect(folder_path):
    """
    遍历文件夹，解析所有 .cma 文件。
    返回 (all_records, all_keys):
      - all_records: 列表，元素形如 {"__filename": 文件名, "A": "xxx", "B": "yyy", ...}
      - all_keys: 除了 __filename 以外出现过的所有键的集合
    """
    all_records = []
    all_keys = set()  # 用于收集所有出现的键
    
    for root, _, files in os.walk(folder_path):
        for f in files:
            if f.lower().endswith(".cma"):
                full_path = os.path.join(root, f)
                blocks = parse_cma_file(full_path)
                
                # 对每个 *LINE block 提取键值对
                for block in blocks:
                    kv_dict = extract_key_values(block)
                    if kv_dict: 
                        # 记录文件名
                        kv_dict["__filename"] = f
                        # 加入全局
                        all_records.append(kv_dict)
                        
                        # 更新全局key集合（去掉特殊的 "__filename"）
                        for k in kv_dict.keys():
                            if k != "__filename":
                                all_keys.add(k)
    
    return all_records, all_keys

def write_to_excel(records, keys, output_excel="output.xlsx"):
    """
    将数据写入Excel：
      - 第一行是 "文件名" + 所有keys（可排序后再写，保证列顺序一致）。
      - 之后每行对应一个块的数据。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "CMA数据"

    # 可以对 keys 排序，以免列顺序混乱
    sorted_keys = sorted(keys)
    
    # 写表头
    ws.cell(row=1, column=1, value="文件名")
    col_index = 2
    for k in sorted_keys:
        ws.cell(row=1, column=col_index, value=k)
        col_index += 1

    # 写每条记录
    current_row = 2
    for rec in records:
        # 第一列写文件名
        ws.cell(row=current_row, column=1, value=rec["__filename"])
        
        # 之后写对应的键值
        col_index = 2
        for k in sorted_keys:
            val = rec.get(k, "")  # 如果没这个键，就空着
            ws.cell(row=current_row, column=col_index, value=val)
            col_index += 1
        
        current_row += 1

    # 保存Excel
    wb.save(output_excel)
    print(f"数据已写入 {output_excel} 文件。")

if __name__ == "__main__":
    # 通过 tkinter 打开文件夹选择对话框
    root = tk.Tk()
    root.withdraw()
    
    folder_to_search = filedialog.askdirectory(title='请选择要扫描的文件夹')
    if folder_to_search:
        # 1) 扫描并收集所有CMA数据
        all_records, all_keys = scan_cma_and_collect(folder_to_search)
        
        # 2) 写入Excel（shift_jis 仅用来读文件，Excel中一般用xlsx本身的编码）
        output_file = "cma_data.xlsx"
        write_to_excel(all_records, all_keys, output_file)
    else:
        print("未选择任何文件夹。")