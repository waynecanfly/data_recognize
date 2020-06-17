import pandas as pd
import xlrd

def get_continuous_list(input_list, yi):
    '''
    获取连续元素列 -> generator
    :param input_list : 一个列表
    ：param yi: 列表在文件中的行
    ：return ：连续元素列表生成器
    '''
    options = []
    for xi, item in enumerate(input_list):
        if item != '':
            item_i = (item, xi ,yi)
            options.append(item_i)
        elif options:
            yield options
            options = []
    if options:
        yield options


def get_excel_info(file_path, sheet_name):
    '''
    :param file_path: 文件路径
    ：param sheet_name : excel文件的sheet_name
    :return :worksheet 对象，有效行数
    '''
    workxls = xlrd.open_workbook(file_path)
    worksheet = workxls.sheet_by_name(sheet_name)
    row = worksheet.nrows
    return worksheet, row


def is_column_line(elem, worksheet):
    '''
    非excel首行数据判断是否是列名所在行
    '''
    if len(elem) == 1:
        pass
    else:
        v, start_xi, yi = elem[0]
        v, end_xi, yi = elem[-1]
        last_row = worksheet.row_values(yi-1)
        last_row_val = last_row[start_xi: end_xi + 1]
        if list(set(last_row_val)) == ['']:
            return True
        else:
            return False


def get_columns(worksheet, row):
    '''
    :param worksheet:worksheet对象
    :param row:sheet中有效行
    ：return 列表生成器
    '''
    for yi in range(row):
        if yi == 0:
            rowdata = worksheet.row_values(yi)
            if len(set(rowdata)) == 1:
                pass
            else:
                elems = get_continuous_list(rowdata, yi)
                for elem in elems:
                    yield elem
        else:
            rowdata = worksheet.row_values(yi)
            if len(set(rowdata)) == 1:
                pass
            else:
                elems = get_continuous_list(rowdata, yi)
                for elem in elems:
                    yield elem


def get_table_df(column, start_xi, end_xi, yi, worksheet, row_nums, sheet_name):
    '''
    :param column:
    :return: 单张数据表的df， 和`sheet` + `column_name(split by '|')`
    '''
    columns = (v for v, xi, yi in column)
    df = pd.DataFrame(columns=columns)
    for raw_num in range(yi, row_nums -1):
        next_raw = raw_num + 1
        next_raw_list = worksheet.row_values(next_raw)
        next_raw_value = next_raw_list[start_xi: end_xi + 1]
        if list(set(next_raw_value)) == ['']:
            break
        else:
            df.loc[len(df)] = next_raw_value
    df_info = sheet_name + ': ' + '|'.join(columns)
    return df, df_info


def get_tables_dataframe(columns, worksheet, row_nums, sheet_name):
    dfs =[]
    df_infos = []
    for column in columns:
        v, start_xi, yi = column[0]
        v, end_xi, yi = column[-1]
        df, df_info = get_table_df(column, start_xi, end_xi, yi, worksheet, row_nums, sheet_name)
        dfs.append(df)
        df_infos.append(df_info)
    return dfs, df_infos


def main():
    file_path = ''
    sheet_name = ''
    worksheet, row_nums = get_excel_info(file_path, sheet_name)
    columns = get_columns(worksheet, row_nums)
    dfs, df_infos = get_tables_dataframe(columns, worksheet, row_nums, sheet_name)
    return dfs, df_infos
