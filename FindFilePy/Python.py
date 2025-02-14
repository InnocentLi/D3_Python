import os
import re
import openpyxl
from openpyxl import Workbook

import tkinter as tk
from tkinter import filedialog

def parse_cma_file(file_path):
    """
    解析单个 .cma 文件，返回其中的所有 *LINE 块。
    每个块会以 (块内容字符串) 的形式返回。
    """
    results = []
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 清洗行：去掉首尾空格和换行
    lines = [line.strip() for line in lines]
    
    current_block = []
    collecting_block = False  # 标识是否在收集当前 *LINE 块

    for line in lines:
        # 检测是否为 *LINE 开头
        if line.startswith("*LINE"):
            # 如果之前有未结束的块，可以先存起来
            if collecting_block and current_block:
                # 上一个块结束
                combined_line = " ".join(current_block)
                results.append(combined_line)
                current_block = []
            
            # 开始新的块收集
            collecting_block = True
            current_block.append(line)  # 把这行也先存起来

        else:
            if collecting_block:
                current_block.append(line)
                # 如果本行包含 `;`，则表示这个块结束
                if ';' in line:
                    combined_line = " ".join(current_block)
                    results.append(combined_line)
                    
                    # 重新开始等待下一个 *LINE
                    collecting_block = False
                    current_block = []
    
    # 如果文件末尾还有没结束的块，也存一下
    if collecting_block and current_block:
        combined_line = " ".join(current_block)
        results.append(combined_line)
    
    return results

def extract_key_values(block_str):
    """
    给定一个块的字符串，比如：
        "*LINE A=123 B=456 C=XX ;"
    返回 "A=123 B=456 C=XX" 这样的字符串。
    可以根据需要进行更复杂的处理。
    """
    # 1) 去掉 *LINE
    block_str = block_str.replace("*LINE", "")
    # 2) 按分号拆分，取分号前内容
    block_str = block_str.split(";")[0]
    
    # 去掉多余空白
    block_str = block_str.strip()
    
    # 用正则找出所有 "X=Y" 形式的文本
    pattern = r'(\w+\s*=\s*[^=\s]+)'
    matches = re.findall(pattern, block_str)
    
    # 将提取到的部分用空格拼接
    return " ".join(matches)

def cma_to_excel(folder_path, output_excel="output.xlsx"):
    """
    遍历 folder_path 下所有 .cma 文件，
    每当遇到 *LINE，解析 key=value 内容，
    最后将结果写入 Excel。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "CMA数据"

    # 写表头（可根据需求更改）
    ws.cell(row=1, column=1, value="文件名")
    ws.cell(row=1, column=2, value="提取值")

    current_row = 2

    # 遍历文件
    for root, _, files in os.walk(folder_path):
        for f in files:
            if f.lower().endswith(".cma"):
                full_path = os.path.join(root, f)
                blocks = parse_cma_file(full_path)
                
                # 对每个块进行key-value的提取
                for block in blocks:
                    kv_str = extract_key_values(block)
                    
                    # 写入Excel
                    ws.cell(row=current_row, column=1, value=f)     # 文件名
                    ws.cell(row=current_row, column=2, value=kv_str) # 提取的值
                    current_row += 1
    
    wb.save(output_excel)
    print(f"数据已写入 {output_excel} 文件。")

if __name__ == "__main__":
    # 使用 Tkinter 打开文件夹选择对话框
    root = tk.Tk()
    root.withdraw()
    
    folder_to_search = filedialog.askdirectory(title='请选择要扫描的文件夹')
    if folder_to_search:
        output_file = "cma_data.xlsx"  # 也可以自己修改名字或用另一个对话框选择保存位置
        cma_to_excel(folder_to_search, output_file)
    else:
        print("未选择任何文件夹。")