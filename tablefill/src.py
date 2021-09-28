""" 
@project: excelfilling
@file: src 
@time: 2021/09/28
@software: PyCharm
@author: zy7y
功能实现
"""

import json

import xlrd
from faker import Faker
from xlutils.copy import copy
from typing import List, Dict


def write_excel(score: str, data: List[str], path: str, sheet_index: int = 0):
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
            if column["type"] == "faker":
                if parameter := column.get("var", False):
                    result = getattr(Faker("zh-Cn"), column["func"])(**parameter)
                else:
                    result = getattr(Faker("zh-Cn"), column["func"])()
            if column["type"] == "input":
                result = column.get("var", "")

            if (var_first := column.get("varFirst")) is not None:
                result = var_first + str(result)

            if (var_end := column.get("varEnd")) is not None:
                result += str(var_end)

            row.append(str(result))
        yield row


def main(conf_path: str, score_file: str, generate_file: str, number: int = 10):
    """
    主程序入口
    :param conf_path: 配置文件路径
    :param score_file: 源excel文件
    :param generate_file: 生成的文件
    :param number: 生成数据量
    :return:
    """
    with open(conf_path, encoding='utf-8') as f:
        conf = json.loads(f.read())
    write_excel(score_file, data=generate_row(conf, number), path=generate_file)


if __name__ == "__main__":
    main("jzg.json", "学生信息导入模板.xls", "jzg.xls")
