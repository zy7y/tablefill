""" 
@project: excelfilling
@file: src 
@time: 2021/09/28
@software: PyCharm
@author: zy7y
功能实现
"""

import json
from pathlib import Path

import xlrd
from faker import Faker
from xlutils.copy import copy
from typing import List, Dict

fake = Faker("zh-Cn")


def write_excel(score: str, data: List[str], path: str, sheet_index: int):
    """
    追加写入excel
    :param score: 源文件地址
    :param data: 写入数据 一个可迭代对象
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
    for line, info in enumerate(data):
        for index, value in enumerate(info):
            new_worksheet.write(rows_exists + line, index, value)
        new_workbook.save(path)


def generate_row(column_conf: List[Dict[str, str]], num: int = 10):
    """
    生成行数据
    :param column_conf: 列配置
    :param num: 10
    :return:
    """
    for _ in range(num):
        row = []
        for column in column_conf:
            if column.get("type", "faker") == "faker":
                if parameter := column.get("var", False):
                    result = getattr(fake, column["func"])(**parameter)
                else:
                    result = getattr(fake, column["func"])()
            elif column["type"] == "input":
                result = column.get("var", "")
            else:
                raise TypeError("type field must faker or input.")

            if (var_first := column.get("varFirst")) is not None:
                result = var_first + str(result)

            if (var_end := column.get("varEnd")) is not None:
                result += str(var_end)

            row.append(result)
        yield row


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
    with open(conf_path, encoding="utf-8") as f:
        conf = json.loads(f.read())
    write_excel(
        score_file,
        data=generate_row(conf, number),
        path=generate_file,
        sheet_index=sheet_index,
    )


__all__ = [
    "main"
]

if __name__ == "__main__":
    main("../examples/demo.json", "../examples/demo.xlsx", "../examples/fakedemo.xls")