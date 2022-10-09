import os
import time

# from pandas.tests.io.excel.test_xlrd import xlrd
import pandas as pd
from pandas.tests.io.excel.test_xlrd import xlrd


class MakePandas():

    def append_excel(self, df, content_list, file_name):
        """
        excel文件中追加内容
        :return:
        content_list:待追加的内容列表
        """
        ds = pd.DataFrame(content_list)
        df = df.append(ds, ignore_index=True)
        excel_name = file_name
        excel_path = os.path.dirname(os.path.abspath(__file__)) + excel_name
        df.to_excel(excel_path, index=False, header=False)

    def remove_row(self, df, row_list):
        """
        excel删除指定列
        :param df:
        :param row_list:
        :return:
        """
        df = df.drop(columns=row_list)
        return df

    def create_excel(self):
        """
        创建excel文件
        :return:
        """
        file_name = '/' + str(time.time()) + 'demo.xlsx'
        file_path = os.path.dirname(os.path.abspath(__file__)) + file_name
        df = pd.DataFrame()
        df.to_excel(file_path, index=False)
        return file_name


txt_name = r'整理前报表.xlsx'
workbook = xlrd.open_workbook(txt_name)
m = MakePandas()
file_name_2 = m.create_excel()
file_name_1 = file_name_2[1:]

result_txt_name = str(time.time())+'.txt'


# print(file_name_2)
for ii in range(3, int(workbook.nsheets)):
    # 读取表格
    df = pd.read_excel(txt_name, sheet_name=ii)

    # 获取最大行数
    max_len = len(df)
    # 当前操作表的时间
    table_time = df.iat[1, 1]
    print(table_time + '处理完毕')

    for i in range(max_len):
        # var参数是获取的付款方式这一列，用来判断临牌
        var = df.iat[i, 5]
        # print(var, type(str(var)), len(str(var)))

        # var2参数是用来获取备注这一列，用来判断上牌
        var2 = df.iat[i, 10]
        # 下面是用来判断临牌的语句
        if var == '月结郝伟伟':
            # print('月结====临牌郝伟伟', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(1)
            a.append('临牌郝伟伟')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 车主
            a.append(df.iat[i, 4])  # 车牌号
            a.append(20)  # 200金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '月结公司':
            # print('月结====月结公司', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(2)
            a.append('临牌公司')
            a.append(df.iat[1, 1])  # 日期
            a.append(df.iat[i, 3])  # 客户
            a.append(df.iat[i, 4])  # 车牌号
            a.append(df.iat[i, 9])  # 车架号
            a.append(df.iat[i, 10])  # 销售姓名
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '月结荣威':
            # print('月结====临牌荣威', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(3)
            a.append('临牌荣威')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 4])  # 车牌号
            a.append('荣威牌')  # 型号
            a.append(df.iat[i, 3])  # 车主
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '月结凡德':
            # print('月结====临牌凡德', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(4)
            a.append('临牌凡德')
            a.append(df.iat[1, 1])  # 日期
            a.append(df.iat[i, 3])  # 车主
            a.append(df.iat[i, 4])  # 车牌号
            a.append(350)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '月结五菱':
            # print('月结====临牌五菱', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(5)
            a.append('临牌五菱')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 4])  # 车牌号
            a.append('五菱牌')  # 车型
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '月结红旗':
            # print('月结====临牌红旗', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(6)
            a.append('临牌红旗')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(20)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)

        # 下面是用来判断上牌的语句
        elif var2 == '月结五菱':
            # print('上牌====月结五菱', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(7)
            a.append('上牌五菱')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 4])  # 车牌号
            a.append('五菱牌')  # 车型
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结凡德':
            # print('上牌====月结凡德', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(8)
            a.append('上牌凡德')
            a.append(df.iat[1, 1])  # 日期
            a.append(df.iat[i, 4])  # 车牌号
            a.append(df.iat[i, 3])  # 车主
            a.append(350)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结荣威':
            # print('上牌====月结荣威', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(9)
            a.append('上牌荣威')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 4])  # 车牌号
            a.append('荣威牌')  # 型号
            a.append(df.iat[i, 3])  # 车主
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结郝伟伟' or var2 == '月结郝':
            # print('上牌====月结郝伟伟', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(10)
            a.append('上牌郝伟伟')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 4])  # 车牌号
            a.append(df.iat[i, 3])  # 车主
            a.append(200)  # 200金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结红旗':
            # print('上牌====月结红旗', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(11)
            a.append('上牌红旗')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)

        # elif var2 == '月结红旗张':
        #     # print('上牌====月结红旗张', i - 1)
        #     m = MakePandas()
        #     df2 = pd.read_excel(file_name_1, header=None)
        #     a = []
        #     b = []
        #     a.append(11)
        #     a.append('上牌红旗')
        #     a.append(df.iat[1, 1])  # 时间
        #     a.append(df.iat[i, 3])  # 客户姓名
        #     a.append(df.iat[i, 4])  # 车牌号
        #     a.append(200)  # 金额
        #     a.append(df.iat[i, 9])  # 车架号
        #     b.append(a)
        #     m.append_excel(df2, b, file_name_2)

        elif var2 == '月结公司':
            # print('上牌====月结公司', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(12)
            a.append('上牌公司')
            a.append(df.iat[1, 1])  # 日期
            a.append(df.iat[i, 3])  # 客户
            a.append(df.iat[i, 4])  # 车牌号
            a.append(df.iat[i, 9])  # 车架号
            a.append(df.iat[i, 10])  # 销售姓名
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '月结红旗张' or var == '月结红旗潘':
            # print('临牌====临牌月结红旗张潘', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(13)
            a.append('月结临牌红旗张潘')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(20)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var == '免费':
            # print('免费', table_time, var2, var, i + 2)
            with open(result_txt_name, 'a', encoding='utf8')as file:
                file.write('免费'+str(table_time)+str(var2)+str(var)+str(i + 2)+'\n')
        elif var2 == '月结红旗张' or var2 == '月结红旗潘':
            # print('上牌====临牌月结红旗张潘', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(14)
            a.append('上牌月结红旗张潘')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结裴秀':
            # print('上牌====月结裴秀', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(15)
            a.append('上牌月结裴秀')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '散客':
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(16)
            a.append('散客')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结岳宏泰':
            # print('上牌====月结岳宏泰', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(17)
            a.append('上牌月结岳宏泰')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结张国良':
            # print('上牌====月结岳宏泰', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(18)
            a.append('上牌月结张国良')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif var2 == '月结奔驰':
            # print('上牌====月结岳宏泰', i - 1)
            m = MakePandas()
            df2 = pd.read_excel(file_name_1, header=None)
            a = []
            b = []
            a.append(19)
            a.append('上牌月结奔驰')
            a.append(df.iat[1, 1])  # 时间
            a.append(df.iat[i, 3])  # 客户姓名
            a.append(df.iat[i, 4])  # 车牌号
            a.append(200)  # 金额
            a.append(df.iat[i, 9])  # 车架号
            b.append(a)
            m.append_excel(df2, b, file_name_2)
        elif str(var) == 'nan' or str(var2) == 'nan':
            continue
        elif var == '月结余' or var == '微信' or var == '收款方式' or var == '流水号':
            continue
        else:
            # print('好奇怪====我也没见过', table_time, var2, var, i + 2)
            with open(result_txt_name, 'a', encoding='utf8')as file:
                file.write('好奇怪====我也没见过'+str(table_time)+str(var2)+str(var)+str(i + 2)+'\n')

# if __name__ == '__main__':
#     MakePandas().create_excel()
    # excel_name = "/demo.xlsx"
    # excel_path = os.path.dirname(os.path.abspath(__file__)) + excel_name
    #
    # m = MakePandas()
    # df = pd.read_excel(excel_path, header=None)
    #
    # b = []
    # for i in range(1, 10):
    #     a = []
    #     a.append(i)
    #     a.append(i * 2)
    #     b.append(a)
    #
    # df = m.append_excel(df, b)
    # # print(df)
