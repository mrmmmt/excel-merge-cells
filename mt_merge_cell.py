import pandas as pd
import xlsxwriter


def excel_merge_cells(df, save_name, key_cols=[], merge_cols=[]):
    '''key_cols：用于判断是否需要合并的key_cols列表
       merge_cols：用于指明哪些列上的单元格需要被合并的列表'''
    self_copy = df.copy(deep=True)
    line_cn = self_copy.shape[0]
    self_copy.index = list(range(line_cn))
    cols = list(self_copy.columns)
    self_copy['temp_col'] = 1
    if all([v in cols for v in key_cols]) == False:   # 校验key_cols中各元素 是否都包含与对象的列
        raise ValueError("key_cols is not completely include object's columns")
    if all([v in cols for v in merge_cols]) == False: # 校验merge_cols中各元素 是否都包含与对象的列 
        raise ValueError("merge_cols is not completely include object's columns")

    wb = xlsxwriter.Workbook(save_name)
    worksheet = wb.add_worksheet()
    format_top = wb.add_format({'border':1, 'bold':True, 'text_wrap':True})
    format_other = wb.add_format({'border':1,'valign':'vcenter'})

    for i, value in enumerate(cols):  # 写表头
        worksheet.write(0, i, value, format_top)

    if key_cols == []:       # 如果key_cols 参数不传值，则无需合并，RN和CN为辅助列
        self_copy['CN'] = 1  # 判断CN大于1的，该分组需要合并，否则该分组（行）无需合并（CN=1说明这个分组数据行是唯一的，无需合并）
        self_copy['RN'] = 1  # RN为需要合并一组中第几行，CN=1，RN=1；CN=5，RN=1,...5
    else:
        self_copy['CN'] = self_copy.groupby(key_cols, as_index=False)['temp_col'].rank(method='max')['temp_col']    # method='max'，对整个组使用最大排名
        self_copy['RN'] = self_copy.groupby(key_cols, as_index=False)['temp_col'].rank(method='first')['temp_col']  # method='first'，按照值在数据中出现的次序分配排名
    
    for i in range(line_cn):
        if self_copy.loc[i, 'CN'] > 1:
            for j, col in enumerate(cols):
                if col in (merge_cols):
                    if self_copy.loc[i, 'RN'] == 1: # 合并写第一个单元格，下一个第一个将不再写
                        worksheet.merge_range(i+1, j, i+int(self_copy.loc[i, 'CN']), j, self_copy.loc[i, col], format_other)
                        '''合并 开始行，开始列，结束行，结束列，值，格式'''
                        '''因为已经写了表头所以从i+1行开始写'''
                    else:
                        pass
                else:
                    worksheet.write(i+1, j, self_copy.loc[i, col], format_other)
        else:
            for j, col in enumerate(cols):
                worksheet.write(i+1, j, self_copy.loc[i, col], format_other)

    wb.close()


if __name__ == '__main__':
    df = pd.DataFrame({'A':[1,2,2,2,3,3],'B':['a','a','a','s','e','f'],'C':['f','c','x','f','w','e'],'D':[1,1,1,1,1,1]})
    print(df)
    excel_merge_cells(df, '000_1.xlsx', ['A'], ['A', 'B'])