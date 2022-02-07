[![Downloads](https://pepy.tech/badge/tablefill)](https://pepy.tech/project/tablefill)
# 背景
在测试Web后台管理系统项目时，导入数据是个高频出现的功能，[tablefill](https://github.com/zy7y/tablefill)主要完成根据配置文件对模板进行填充数据

# 使用
**安装**
```shell
pip install tablefill
```
**配置列数据类型**
```json5
[
  {
    "type": "faker", // 可选值 faker(默认值,可不写type这个字段)、input 会直接读取var 的值 由自己设置
    "func": "name", // 对应的是 Faker 生成虚拟数据的那些方法名 https://faker.readthedocs.io/en/master/providers.html
    "var": null, // 没有参数时可以不写该字段, 当type 为faker时 这部分会被作为func 对应函数名的入参
    "varFirst": "前", // 如果不需要可以不写该字段, 会在 var 这个 参数 前面 加上 内容
    "varEnd": "后" // 如果不需要可以不写该字段, 会在 var 这个 参数 后面 加上 内容
  }
]
```
**api参考[faker](https://faker.readthedocs.io/en/stable/providers.html)**
phone_number: 生成手机号
random_element: 列表中随机元素
name: 随机名称
ssn: 身份证号
date: 随机日期


*示例*
```json
[
  {
    "type": "input",
    "var": "这列我输入"
  },
  {
    "func": "phone_number"
  },
  {
    "func": "random_int",
    "var": {
      "min": 10,
      "max": 21
    },
    "varFirst": "编号",
    "varEnd": "班"
  },
  {
    "func": "random_element",
    "var": {
      "elements": ["小学", "高中", "初中"],
    }
  }
]
```
**导入模板文件**
> 需要是xlsx/xls文件
[![4h3G3F.md.png](https://z3.ax1x.com/2021/09/29/4h3G3F.md.png)](https://imgtu.com/i/4h3G3F)

**执行命令**
```shell
# --num 可选参数 默认 10条 ，这里就是30条
fill generate 配置文件 模板文件 生成文件名 --num 30 

fill generate "E:\coding\tablefill\examples\demo.json" "E:\coding\tablefill\examples\demo.xlsx"  demo.xls
```

**填充数据后的文件**
[![4h8FbR.md.png](https://z3.ax1x.com/2021/09/29/4h8FbR.md.png)](https://imgtu.com/i/4h8FbR)

**help**
```shell
fill --help
```

