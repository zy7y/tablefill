""" 
@project: excelfilling
@file: cli 
@time: 2021/09/28
@software: PyCharm
@author: zy7y
命令行扩展
"""
from pathlib import Path
from typing import Optional

from typer import Typer
from typer import echo
from typer import launch
from typer import confirm

from .src import main


app = Typer(help="- Excel表格模板数据填充工具 -")


@app.command("docs")
def visit_docs():
    """
    访问tablefill文档
    :return:
    """
    visit = confirm("即将访问tablefill文档，确认?")
    if visit:
        launch("https://gitee.com/zy7y/tablefill.git")


@app.command()
def files(path: Path, suffix: Optional[str] = ".json"):
    """
    列出绝对路径下所有 后缀为 xx 文件的绝对路径
    :param path: 需要查找的目录绝对路径
    :param suffix: 查找文件的后缀
    :return:

    >>> file_list(Path(__file__).absolute().parent)

    """
    p = Path(path)
    for child in p.glob(f"**/*{suffix}"):
        echo(child.absolute())


@app.command()
def generate(
    config: str,
    template: str,
    file: str,
    num: Optional[int] = 10,
    index: Optional[int] = 0,
):
    """
    读取行配置文件填充 excel 文件
    :param config: 列配置文件
    :param template: 待填充的excel模板
    :param file: 填充后的存储位置
    :param num: 填充总行数 默认 10条
    :param index: 填充到指定sheet页 默认 0 即 第一个
    :return:
    """
    main(config, template, file, num, index)


if __name__ == "__main__":
    app()
