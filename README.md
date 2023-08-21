# easy excel tool

在实际工作中，沉淀的一些简易、好用的excel数据处理工具，减少重复代码与文件冗余，希望一样能为使用者带来益处。如果您也想贡献好的代码片段，请将代码以及描述，通过邮箱（ [xinkonghan@gmail.com](mailto:hanxinkong<xinkonghan@gmail.com>)
）发送给我。代码格式是遵循自我主观，如存在不足敬请指出！

## 安装

```shell
pip install easy-excel-tool
```

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

## 链接

Github：https://github.com/hanxinkong/easy-excel-tool

在线文档：https://easy-excel-tool.xink.top/

## 注明