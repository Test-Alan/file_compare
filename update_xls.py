
import xlrd
from xlutils.copy import copy
import csv
import os
from file_comparison import *

'''
   需求说明：
    A表中有几十个用户信息，在B目录中根据A表的用户数量创建了相同数量的表格文件，并以用户名命名。
    需要把A表中的数据根据用户名去更新B目录中的所有表格中某个单元的值。
    实现过程：
    1.获取A表中的所有用户名，和要更新的数据。
    2.获取B目录中的所有文件名
    3.根据用户名去匹配文件名，如果用户名和文件名相同就更新表格内的值。
'''

# 获取A表的数据
def open_csv_file(file_dir):
    '''
    读取csv文件
    :return:
    '''
    data = csv.reader(open(file_dir))
    return data

# 更新表的某个单元格
def excelUpdate(excel,data,save_path):
    '''
    更新已有excel, 是copy一份再做更新
    :param excel:表格路径，data:更新的值 save_path:更新后保存的路径
    :return:
    '''
    wb = xlrd.open_workbook(excel)
    newb = copy(wb)  # 类型为worksheet 无nrows方法
    wbsheet = newb.get_sheet(0)
    # 要更新的单元格和更新的值
    wbsheet.write(4, 7, data)
    # 保存的路径
    newb.save(save_path)

# 获取B目录所有文件名
def get_files_name(file_path):

    # 获取要修改标的所有文件名
    files_name = get_remove_suffix_filename(file_path)
    return files_name

# 根据A表的用户名去匹配B目录的文件名
def file_matching():

    # 获取A表的数据
    datas = open_csv_file(r'D:\提前还款金额2.csv')
    # 获取文件名
    names = get_files_name(r'E:\代偿用户对比\表格有的合同没有的')
    # 通过用户名循环匹配文件名
    for data in datas:
        # 获取csv文件的用户名
        username = data[0]
        for file_name in names:
            # 判断cvs文件的用户是否等于被修改目录文件的用户名，如果是的话就把要修改的数据传给excelUpdate方法修改
            if username == file_name:
                lixi = data[2].split(':')
                user_lixi = lixi[-1]
                print(username, user_lixi)
                path = r'E:\代偿用户对比\表格有的合同没有的' + '\\' + username + '.xls'
                save_path = 'F:\\' + username + '.xls'
                excelUpdate(path, user_lixi,save_path)

if __name__ == '__main__':
    file_matching()
