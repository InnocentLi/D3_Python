import os
import json
import asyncio
from aiofiles import open as aio_open

# 用于限制并发数量，避免递归层数过多时任务爆炸式增长
# 根据实际情况调整此数值。数值过大会增加系统负担，过小则并发度不足。
SEM = asyncio.Semaphore(100)

async def list_files_async(directory: str) -> list[str]:
    """
    异步递归遍历目录，返回文件路径列表。

    :param directory: 要遍历的目录路径
    :return: 文件路径列表
    """
    file_list = []

    try:
        # 用 to_thread+scandir 提升 I/O 性能，并异步获取目录项
        async with SEM:
            entries = await asyncio.to_thread(os.scandir, directory)
    except PermissionError:
        # 无权限时跳过此目录
        print(f"权限不足，跳过目录: {directory}")
        return file_list
    except Exception as e:
        print(f"访问目录时发生错误: {directory}\n错误信息: {e}")
        return file_list

    # 准备并行处理子目录任务
    tasks = []
    for entry in entries:
        if entry.is_file():
            file_list.append(entry.path)
        elif entry.is_dir():
            tasks.append(list_files_async(entry.path))

    # 并行处理所有子目录
    if tasks:
        results = await asyncio.gather(*tasks, return_exceptions=True)
        for r in results:
            # 如果有子任务出错，这里可以处理异常或跳过
            if isinstance(r, Exception):
                print(f"子任务发生异常: {r}")
            elif isinstance(r, list):
                file_list.extend(r)

    return file_list

async def save_to_json(data: list[str], output_file: str):
    """
    异步保存数据到 JSON 文件。

    :param data: 要保存的数据（文件列表）
    :param output_file: 输出 JSON 文件路径
    """
    try:
        async with aio_open(output_file, mode='w', encoding='utf-8') as f:
            json_data = json.dumps(data, indent=4, ensure_ascii=False)
            await f.write(json_data)
        print(f"文件列表已保存到: {output_file}")
    except Exception as e:
        print(f"保存 JSON 文件时发生错误: {e}")

async def main():
    """
    主异步任务：
    1. 从用户处获取要遍历的目录和输出文件路径
    2. 递归遍历文件
    3. 保存结果到指定的 JSON 文件
    """
    directory = input("请输入要遍历的目录路径: ").strip()
    output_file = input("请输入输出 JSON 文件路径: ").strip()

    print(f"开始遍历目录: {directory}\n")
    file_list = await list_files_async(directory)
    await save_to_json(file_list, output_file)

if __name__ == "__main__":
    asyncio.run(main())
