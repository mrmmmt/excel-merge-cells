# excel-merge-cells 利用 Python 实现 Excel 合并单元格
利用 Python 进行 Excel 文件读取并分析多数采用 Pandas 模块进行，所以直接将方法定义为一个名为 excel_merge_cells 的函数，函数需要传入的参数如下：

- df：需要进行合并单元格处理的DataFrame
- save_name：处理结果需要保存的路径及文件名
- key_cols：用于判断是否需要合并的key_cols列表
- merge_cols：用于指明哪些列上的单元格需要被合并的列表
