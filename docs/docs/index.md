一些简易、好用的excel数据处理工具，减少重复代码与文件冗余

----
**文档地址：** <a href="/https://easy-excel-tool.xink.top/" target="_blank">https://easy-excel-tool.xink.top/ </a>

**PyPi地址：
** <a href="https://pypi.org/project/easy-excel-tool" target="_blank">https://pypi.org/project/easy-excel-tool </a>

**GitHub地址：** [https://github.com/hanxinkong/easy-excel-tool](https://github.com/hanxinkong/easy-excel-tool)

----

## 安装

<div class="termy">

```console
pip install easy-excel-tool
```

</div>

## 主要功能

- `excel`
    - `create_excel` 创建一个空白的excel文件（可指定sheet名）
    - `add_sheet` 对指定excel文件新增sheet页（可指定sheet名）
    - `get_sheet_name` 获取excel文件所有sheet名
    - `remove_sheet` 删除excel文件中指定sheet页
    - `write_excel` 对excel文件写入或追加数据（可指定填充列，保留头部）

## 简单使用

```python
from easy_excel_tool import Excel

excel = Excel('./test.xlsx')
# excel.create_excel(inplace=True)
# excel.add_sheet(sheet_name=[
#     'donation_information',
# ])
data = [
    {'a': 1, 'b': 5, 'c': '88'},
    {'a': 7, 'b': 9, 'c': 66},
]
excel.write_excel(
    data,
    mode='w+',
    sheet_name='Sheet1',
    fill_column={'fill_column': 'fill'},
    header=True,
    inplace=False,
)
```

## 依赖

内置依赖

- `os` The os module in Python provides a way to use operating system dependent functionalities.
- `typing` Type Hints for Python.

第三方依赖

- `pandas` Pandas is a popular open-source data analysis and manipulation library for Python.
- `openpyxl` openpyxl is a Python library for reading and writing Excel files, specifically the newer .xlsx format.

_注：依赖顺序排名不分先后_

## 许可证

该项目根据 **MIT** 许可条款获得许可.

## 注明
