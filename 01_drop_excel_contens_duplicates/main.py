# _*_ coding:utf-8 _*_
# Author: zizle
# -----------------------------------
# function:  使用pandas读取excel文件内容与数据库中数据做对比，合并后去重，保存至新的excel文件
# environments:
# 1 python 3.6.3
# 2 numpy 1.18.3
# 3 pandas  1.0.3
# 4 python-dateutil 2.8.1
# 5 pytz 2020.1
# 6 six 1.14.0
# -----------------------------------
import pandas as pd


def read_excel_contents():
    file_contents = xlrd.open_workbook(filename='muban.xlsx')
    # table_data = file_contents.sheets()[0]
    # 导入名称为“值班信息”的表
    table_data = file_contents.sheet_by_name('值班信息')
    # 检查sheet1是否导入完毕
    status = file_contents.sheet_loaded('值班信息')
    if not status:
        print('文件数据导入失败')
    # 读取第一行数据
    first_row = table_data.row_values(0)
    # 格式判断
    if first_row != [
        "日期", "信息内容", "备注"
    ]:
        print("表格格式有误,请修正.")
    nrows = table_data.nrows
    # ncols = table_data.ncols
    # print("行数：", nrows, "列数：", ncols)
    # 获取数据
    ready_to_save = list()  # 准备保存的数据集
    start_data_in = False
    # 组织数据写入数据库
    message = "表格列数据类型有误,请检查后上传."
    try:
        for row in range(nrows):
            row_content = table_data.row_values(row)
            # 找到需要开始上传的数据
            if str(row_content[0]).strip() == 'start':
                start_data_in = True
                continue  # 继续下一行
            if str(row_content[0]).strip() == 'end':
                start_data_in = False
                continue
            if start_data_in:
                record_row = list()  # 每行记录
                # 转换数据类型
                try:
                    record_row.append(xlrd.xldate_as_datetime(row_content[0], 0))
                except Exception as e:
                    message = "第一列【日期】请使用日期格式上传."
                    raise ValueError(e)
                record_row.append(str(row_content[1]))
                record_row.append(str(row_content[2]))
                ready_to_save.append(record_row)
    except Exception as e:
        import traceback
        traceback.print_exc()
        print('读取数据错误:', e)

    file_df = pd.DataFrame(ready_to_save)
    file_df.columns = ['custom_time', 'content', 'note']
    return file_df


def read_db_contents():
    """
    读取已有的数据集
    :return: DataFrame
    """
    db_connection = MySQLConnection()
    cursor = db_connection.get_cursor()
    query_statement = "SELECT * FROM `onduty_message` WHERE `author_id`=%d" % user_id
    cursor.execute(query_statement)
    db_connection.close()
    result = cursor.fetchall()
    # 把字典列表转成DataFrame
    exist_df = pd.DataFrame(result)

    split_df = exist_df[['custom_time', 'content', 'note']]
    return split_df


def concat_data_frame(old_df, new_df):
    return pd.concat([old_df, new_df])


def drop_duplicates(old_df, new_df):
    """
    新增数据根据已有的去重
    :return: DataFrame
    """
    old_df['custom_time'] = pd.to_datetime(old_df['custom_time'], format='%Y-%m-%d')
    new_df['custom_time'] = pd.to_datetime(new_df['custom_time'], format='%Y-%m-%d')
    print('已有数据大小:', old_df.shape)
    print('===========已有数据进行去重==============')
    old_df = old_df.drop_duplicates(subset=['custom_time', 'content'], keep='first', inplace=False)
    print('===========已有数据进行去重完毕!==============')
    print('已有数据大小:', old_df.shape)
    print('新数据大小:', new_df.shape)
    concat_df = concat_data_frame(old_df, new_df)
    print('合并后数据大小:', concat_df.shape)

    finally_df = concat_df.drop_duplicates(subset=['custom_time', 'content'], keep=False, inplace=False)
    print("去重后数据大小:", finally_df.shape)
    if finally_df.empty:
        print('去重后数据为空')
    else:
        print('去重后数据不为空')
        print('去重后数据为:\n', finally_df)
        # 时间处理
        # time_format = finally_df['custom_time'].apply(lambda x: x.strftime('%Y-%m-%d'))
        # finally_df['custom_time'] = time_format
        finally_list = finally_df.values.tolist()
        print('转为列表为:\n', finally_list)


def save_new_excel():
    db_connection = MySQLConnection()
    cursor = db_connection.get_cursor()
    query_statement = "SELECT * FROM `onduty_message` WHERE `author_id`=43;"
    cursor.execute(query_statement)
    db_connection.close()
    result = cursor.fetchall()
    # 把字典列表转成DataFrame
    exist_df = pd.DataFrame(result)
    exist_df['custom_time'] = exist_df['custom_time'].apply(lambda x: x.strftime('%Y-%m-%d'))
    # split_df = exist_df[['custom_time', 'content', 'note']]
    exist_df.to_excel(
        excel_writer='new_excel.xlsx',
        index=False,
        sheet_name="值班信息",
        columns=['custom_time', 'note']
    )
    return exist_df


if __name__ == '__main__':
    pass

