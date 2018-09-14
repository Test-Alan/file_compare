import xlrd
import docx
import re
from win32com import client as wc
import os
from time import sleep


'''
    需求说明：
    需要检查100多份合同里面的数据是否填错，根据的数据是从数据库查询出来然后放在Excel表格中的。
    如果手动去对比需要一个一个查看很麻烦，所以写个简单的自动化脚本。
    实现步骤：
    1.把需要检查的数据从Excel表格中提取出来，放在list列表中。
    2.把合同的数据根据正则匹配提取出来，也放在list列表中。
    3.然后对比两个list是否相同
    4.如果不同就输出把不同的找出来。
'''
class OpenExcel():
    '''
    读取excel文件
    '''
    #data = xlrd.open_workbook(r'E:\代偿用户对比\对比文件\陈晔.xls')
    def __init__(self, path):
        # 打开excel文件
        self.data = xlrd.open_workbook(path)
        self.sheet = self.data.sheet_by_index(0)
    # 获取单元格的值
    def get_table_value(self, *args):
        # 获取读入的文件的

        a, b = args[0]
        data = self.sheet.cell(a, b).value
        return data


class OpenDocx():
    '''
    读取docx文件
    '''

    def __init__(self, doc_path, docx_path):
        self.doc_path = doc_path
        self.docx_path = docx_path

    # 将doc转为docx:
    def doc_convert_docx(self):
        # path = r'E:\代偿用户对比\对比文件\委托代偿协议.doc'
        # path1 = r'E:\代偿用户对比\对比文件\委托代偿协议.docx'
        word = wc.Dispatch("Word.Application")
        # 打开被转换的文件
        doc = word.Documents.Open(self.doc_path)
        # 转换后要保存的路径及文件名
        doc.SaveAs(self.docx_path, 12) # 12为docx
        # 关闭文件
        doc.Close()
        # 退出
        word.Quit()

    # 获取文件内容
    def get_docx_file(self):
        # 读取docx文件
        data = docx.Document(self.docx_path)
        t = []
        # 循环读取每一行内容，并把他加入列表中
        for i in data.paragraphs:
            t.append(i.text)
        # 将列表的内容转换成字符串
        s = "".join(t)
        s = s.replace(' ','')
        #print(s)
        # 正则表达式
        r = '甲方：(.*?)身份证号：(.*?)联系地址.*等借款【(\d+)】.*借款合同编号为：【(.*)】.*补充合同编号为【(.*)】.*已偿还借款本金【(.*)】元.*利息【(.*)】元.*服务费【(.*)】元.*未到期债务的借款本金【(.*)】元.*利息【(.*)】元.*服务费【(.*)】.*甲方参照编号为【(.*)】.*和编号为【(.*)】的补充合同中利息（年利率【(.*)】\%）和还款日期的约定，向乙方清偿代偿款及利息（年利率【(.*)】%）'
        # 匹配内容
        p = re.search(r, s)
        user_data2 = list(p.groups())
        return user_data2


def file_name(file_dir):

    for root, dirs, files in os.walk(file_dir):

        # print('当前目录路径',root)  # 当前目录路径
        # print('当前路径下所有子目录',dirs)  # 当前路径下所有子目录
        # print('当前路径下所有非目录子文件',files)  # 当前路径下所有非目录子文件
        return dirs , files

# 对比后保存的结果。
def save_result(data):
    with open(r'D:\委托代偿协议对比结果.txt', 'a') as f:
        f.write(data)

def main():
    # 单元格数据名称
    a = ['用户名:', '身份证号:', '借款金额:', '合同编号:', '补充合同编号:', '已还本金:', '已还利息:', '已还服务费:', '未还本金:', '未还利息:', '未还服务费:', '借款合同:',
         '合同编号:', '合同年利率:', '偿还年利率:']
    cell = [(1, 3), (1, 4), (1, 1), (1, 6), (9, 6), (3, 6), (3, 7), (3, 8), (4, 6), (4, 7), (4, 8), (1, 6), (9, 6),
            (1, 9), (1, 9)]
    path = r'E:\代偿用户对比\对比文件\表格'
    path2 = r'E:\代偿用户对比\对比文件\合同'
    tible_file = file_name(path)[1]
    hetong_file = file_name(path2)[0]
    # 统计执行了几次
    count = 1
    # 统计不一样的有几次
    bt = 1
    for table_username in tible_file:
        username1 = table_username.split('.')
        username1 = username1[0]
        for hetong_username in hetong_file:
            if username1 == hetong_username :
                excel_file = r'E:\代偿用户对比\对比文件\表格' + '\\' + table_username
                doc_file  = r'E:\代偿用户对比\对比文件\合同' + '\\' + username1 + '\\' + '委托代偿协议.doc'
                docx_file = r'E:\代偿用户对比\对比文件\合同' + '\\' + username1 + '\\' + '委托代偿协议.docx'
                # print(excel_file)
                # print(doc_file)
                # print(docx_file)
                try:
                    user_data = []
                    num = len(cell)-1
                    for data in cell:
                        # 提取cell列表相对应的单元格的值
                        d = OpenExcel(excel_file).get_table_value(data)
                        # 判断某个单元格的值是否是0.0如果是的话就转换成0
                        if d == 0.0 or d == '':
                            d = '0'
                        # 判断是否是最后两个数值，如果是的话就乘以100
                        if (num == 1 or num == 0) and type(d)== float:
                            d = d * 100

                        # 判断某个单元格的值是否是小数，如果是的话就保留两位小数四舍五入
                        if type(d) == float:
                            d = round(d, 2)
                            # 最后小数位是否为0如果是的话就把0去掉
                            if int(d) - d == 0:
                                d = int(d)
                            d = str(d)
                        # 最后吧获取到的时候保存到表格
                        user_data.append(d)
                        num -= 1

                    user_data1 = user_data
                    # 修改doc为docx
                    OpenDocx(doc_path=doc_file, docx_path=docx_file).doc_convert_docx()
                    user_data2 = OpenDocx(doc_path=doc_file, docx_path=docx_file).get_docx_file()


                    # print(user_data1)
                    # print(user_data2)

                    num1 = 0
                    d1 = []
                    for i in a:
                        d1.append(i + user_data1[num1])
                        num1 += 1
                    print(d1)

                    num2 = 0
                    d2 = []
                    for i in a:
                        d2.append(i+user_data2[num2])
                        num2 += 1
                    print(d2)
                    jk_money = float(user_data2[2])
                    yh_money = float(user_data2[5])
                    wh_money = float(user_data2[8])
                    sum_money = yh_money + wh_money
                    sum_money = round(sum_money, 2)

                    print('借款金额:%s，相加总金额：%s' % (jk_money, sum_money))

                    if jk_money == sum_money:
                        result = '金额相等'
                        print('金额相等')
                    else:
                        result = '金额不相等'
                        print('金额不相等')
                        # d1_result = list(set(d1).difference(set(d2)))
                        # d2_result = list(set(d2).difference(set(d1)))
                    print(list(set(d1).difference(set(d2))))
                    print(list(set(d2).difference(set(d1))))
                    print(d1 == d2)
                    b = d1==d2
                    if b == False and len(list(set(d1).difference(set(d2)))) > 0 :
                        bt_d1 = ",".join(d1)
                        bt_d2 = ",".join(d2)
                        bt_d1_result = ",".join(list(set(d1).difference(set(d2))))
                        bt_d2_result = ",".join(list(set(d2).difference(set(d1))))
                        save_result(bt_d1 + '\n')
                        save_result(bt_d2 + '\n\n')
                        save_result('不同的数据如下：' + '\n')
                        save_result('表格文件数据：\t'+ bt_d1_result + '\n\n')
                        save_result('委托代偿协议数据：\t' + bt_d2_result + '\n\n')
                        save_result('借款金额:%s，相加总金额：%s' % (jk_money, sum_money) + '\n')
                        save_result(result + '\n\n')
                        save_result('========================================================' + '\n')
                        print('不相同的有%d个' % bt)
                        bt += 1
                    print('========================== %d ================================' % count)
                    count += 1
                    sleep(2)
                except Exception as msg:
                    print(msg)
                    print(username1, '没有对比成功')
main()