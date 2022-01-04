""" 
@project: excelfilling
@file: src 
@time: 2021/09/28
@software: PyCharm
@author: zy7y
功能实现
"""
import json
import time
from concurrent.futures import ThreadPoolExecutor

import xlrd
from faker import Faker
from rich.console import Console
from rich.progress import track
from xlutils.copy import copy
from typing import Dict
from queue import Queue

fake = Faker("zh-Cn")
console = Console()


def write_excel(new_worksheet, que: Queue, rows_exists):
    """实际写excel 的动作"""
    for index, value in enumerate(que.get()):
        new_worksheet.write(rows_exists + que.qsize(), index, value)


def create_file(score: str, path: str, que: Queue, sheet_index: int = 0):
    """
    多线程追加写入excel
    :param score: 源文件地址
    :param que: 需写入数据的队列
    :param path: 追加后的地址
    :param sheet_index: 追加数据到sheet页某个索引
    :return:
    """
    workbook = xlrd.open_workbook(score)
    sheets_name = workbook.sheet_names()
    worksheet = workbook.sheet_by_name(sheets_name[sheet_index])
    rows_exists = worksheet.nrows
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(sheet_index)

    # 线程池
    with ThreadPoolExecutor() as executor:
        for _ in track(range(que.qsize()), description='写入...'):
            future = executor.submit(write_excel, new_worksheet, que, rows_exists)
            # todo 如不加操作会丢失数据
            future.result()

    new_workbook.save(path)


def row_data(row_cfg: Dict[str, str]) -> str:
    """
    根据列配置生成单个列数据
    :param row_cfg: 列配置
    :return: 虚拟数据
    """
    if row_cfg.get("type", "faker") == "faker":
        if parameter := row_cfg.get("var"):
            result = getattr(fake, row_cfg["func"])(**parameter)
        else:
            result = getattr(fake, row_cfg["func"])()
    elif row_cfg["type"] == "input":
        result = row_cfg.get("var", "")
    else:
        raise TypeError("type field must faker or input.")

    if not isinstance(result, str):
        result = str(result)

    if (var_first := row_cfg.get("varFirst")) is not None:
        result = str(var_first) + result

    if (var_end := row_cfg.get("varEnd")) is not None:
        result += str(var_end)
    return result


def generate_rows(cfg_path: str, que: Queue, number: int):
    """
    生成行数据
    :param cfg_path: 配置文件
    :param que: 队列
    :param number: 生成数量
    :return:
    """
    with open(cfg_path, encoding="utf-8") as f:
        conf = json.loads(f.read())

    [que.put([row_data(row_cfg) for row_cfg in conf]) for _ in track(range(number), description='生成...', total=number)]


def main(
        conf_path: str,
        score_file: str,
        generate_file: str,
        number: int = 10,
        sheet_index: int = 0,
):
    """
    主程序入口
    :param conf_path: 配置文件路径
    :param score_file: 源excel文件
    :param generate_file: 生成的文件
    :param number: 生成数据量
    :param sheet_index: 写在第几个sheet页
    :return:
    """
    console.print("欢迎使用tablefill 任务开始~")
    start = time.time()
    queue = Queue()
    generate_rows(conf_path, queue, number)
    create_file(score_file, generate_file, queue, sheet_index)
    console.print(f"您耗费了{round(time.time() - start, 2)}s, 为您向 {generate_file} 中填充了{number}条数据, 快去看看吧")


__all__ = [
    "main"
]

if __name__ == "__main__":
    main("../examples/demo.json", "../examples/demo.xlsx", "../examples/tablefill.xls")