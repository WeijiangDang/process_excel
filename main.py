# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import numpy as np


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def process_excel():
    # 打开要写入的目标excel
    target_excel_path = 'F:\python_proj\doc\填报工具2021.xlsx'
    read_excel_res = pd.read_excel(target_excel_path, sheet_name='2021年招生计划')
    df = pd.DataFrame(read_excel_res)
    # 打印院校代码和专业代码列
    df_codes = df[['院校代码', '专业代码', '院校名称']]
    print("院校代码 专业代码：\n{0}".format(df_codes))
    col_name = df.columns.tolist()
    print("列：\n{0}".format(col_name))
    # 新增列21年院校最低分
    col_name.insert(42, '21年院校最低分')
    # 新增列
    col_name.insert(28, '21年专业最低分')
    print("列：\n{0}".format(col_name))
    df = df.reindex(columns=col_name)
    print(df[['院校代码', '专业代码', '院校名称', '21年专业最低分', '21年院校最低分']])

    # 获取学校的行号
    # target_academy = ['清华大学']
    # index_academy = df[df.loc[:, '院校名称'].isin(target_academy)].index
    # print(index_academy)

    # 打开院校最低分excel
    academy_min_excel_path = 'F:\python_proj\doc\\2017-2021数据整理.xlsx'
    academy_min_excel_res = pd.read_excel(academy_min_excel_path, sheet_name='2017-2021理科')
    # 打印院校代码和专业代码列
    df_academy_min = academy_min_excel_res[['大学', '21年最低分']]
    print("大学和院校最低：\n{0}".format(df_academy_min))
    df_target_academy_min = df['21年院校最低分']
    print("院校最低：\n{0}".format(df_target_academy_min))
    # 遍历所有院校最低分，对于不是NaN的数据，查询学校在主文档中的行号列表，全部设置为该分数
    for index, row in df_academy_min.iterrows():
        if not pd.isnull(row['大学']):
            target_school = [row['大学']]
            index_academy = df[df.loc[:, '院校名称'].isin(target_school)].index
            # 把所有index的最低分值设置为row['21年最低分']
            for i in index_academy:
                df_target_academy_min[i] = row['21年最低分']
    df['21年院校最低分'] = df_target_academy_min

    # 打开21年非专科录取数据 理科
    sci_major_min_excel_path = 'F:\python_proj\doc\\2021年录取数据.XLS'
    sci_major_min_excel_path_res = pd.read_excel(sci_major_min_excel_path, sheet_name='物理')
    # 打印院校代码和专业代码列
    df_sci_min = sci_major_min_excel_path_res[['院校\n编号', '院校名称', '专业\n编号', '专业名称', '投档\n最低分']]
    print("大学和专业最低：\n{0}".format(df_sci_min))
    df_target_major_min = df['21年专业最低分']
    for index, row in df_sci_min.iterrows():
        if not pd.isnull(row['院校\n编号']):
            # 查询主表中院校编号和专业编号对应的行号列表
            index_list = df.index[(df['院校代码'] == row['院校\n编号']) & (df['专业代码'] == row['专业\n编号'])].tolist()
            for i in index_list:
                df_target_major_min[i] = row['投档\n最低分']
    df['21年专业最低分'] = df_target_major_min

    # 打开21年非专科录取数据 文科
    sci_major_min_excel_path = 'F:\python_proj\doc\\2021年录取数据.XLS'
    sci_major_min_excel_path_res = pd.read_excel(sci_major_min_excel_path, sheet_name='历史')
    # 打印院校代码和专业代码列
    df_sci_min = sci_major_min_excel_path_res[['院校\n编号', '院校名称', '专业\n编号', '专业名称', '投档\n最低分']]
    print("大学和专业最低：\n{0}".format(df_sci_min))
    df_target_major_min = df['21年专业最低分']
    for index, row in df_sci_min.iterrows():
        if not pd.isnull(row['院校\n编号']):
            # 查询主表中院校编号和专业编号对应的行号列表
            index_list = df.index[(df['院校代码'] == row['院校\n编号']) & (df['专业代码'] == row['专业\n编号'])].tolist()
            for i in index_list:
                df_target_major_min[i] = row['投档\n最低分']
    df['21年专业最低分'] = df_target_major_min

    # 打开21年专科录取数据 理科
    sci_major_min_excel_path = 'F:\python_proj\doc\\2021年专科录取数据.XLS'
    # converters显式定义字段为字符串，避免按数字处理不能匹配
    sci_major_min_excel_path_res = pd.read_excel(sci_major_min_excel_path, sheet_name='物理学科类', converters={'院校\n编号': str})
    # 打印院校代码和专业代码列
    df_sci_min = sci_major_min_excel_path_res[['院校\n编号', '院校名称', '专业\n编号', '专业名称', '投档\n最低分']]
    print("大学和专业最低：\n{0}".format(df_sci_min))
    df_target_major_min = df['21年专业最低分']
    for index, row in df_sci_min.iterrows():
        if not pd.isnull(row['院校\n编号']):
            # 查询主表中院校编号和专业编号对应的行号列表
            index_list = df.index[(df['院校代码'] == row['院校\n编号']) & (df['专业代码'] == row['专业\n编号'])].tolist()
            for i in index_list:
                df_target_major_min[i] = row['投档\n最低分']
    df['21年专业最低分'] = df_target_major_min

    # 打开21年专科录取数据 文科
    sci_major_min_excel_path = 'F:\python_proj\doc\\2021年专科录取数据.XLS'
    # converters显式定义字段为字符串，避免按数字处理不能匹配
    sci_major_min_excel_path_res = pd.read_excel(sci_major_min_excel_path, sheet_name='历史类', converters={'院校\n编号': str})
    # 打印院校代码和专业代码列
    df_sci_min = sci_major_min_excel_path_res[['院校\n编号', '院校名称', '专业\n编号', '专业名称', '投档\n最低分']]
    print("大学和专业最低：\n{0}".format(df_sci_min))
    df_target_major_min = df['21年专业最低分']
    for index, row in df_sci_min.iterrows():
        if not pd.isnull(row['院校\n编号']):
            # 查询主表中院校编号和专业编号对应的行号列表
            index_list = df.index[(df['院校代码'] == row['院校\n编号']) & (df['专业代码'] == row['专业\n编号'])].tolist()
            for i in index_list:
                df_target_major_min[i] = row['投档\n最低分']
    df['21年专业最低分'] = df_target_major_min

    # 导出结果到excel
    output_path = 'F:/python_proj/doc/out.xlsx'
    df.to_excel(output_path)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    process_excel()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
