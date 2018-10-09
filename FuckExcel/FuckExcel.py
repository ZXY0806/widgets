# -*- coding: utf-8 -*-
import xlrd,xlwt

# 重复数据生成器
def data_generate1(data_str, rp_num):
    n = 0
    while n < rp_num:
        yield data_str
        n += 1

# 递增数据生成器
def data_generate2(data_str, ic_num, base_left_index, base_right_index):
    left_data_str = data_str[:base_left_index]
    base_num = int(data_str[base_left_index:base_right_index])
    right_data_str = ''
    if base_right_index < len(data_str):
        right_data_str = data_str[base_right_index:]
    n = 0
    while n < ic_num:
        new_data_str = left_data_str + str(base_num+n) + right_data_str
        yield new_data_str
        n += 1
# 重复递增数据生成器
def data_generate3(data_str, rp_num, ic_num, base_left_index, base_right_index):
    left_data_str = data_str[:base_left_index]
    base_num = int(data_str[base_left_index:base_right_index])
    right_data_str = ''
    if base_right_index < len(data_str):
        right_data_str = data_str[base_right_index:]
    n = 0
    while n < ic_num:
        m = 0
        while m < rp_num:
            new_data_str = left_data_str + str(base_num + n) + right_data_str
            yield new_data_str
            m += 1
        n += 1
# 递增重复数据生成器
def data_generate4(data_str, rp_num, ic_num, base_left_index, base_right_index):
    left_data_str = data_str[:base_left_index]
    base_num = int(data_str[base_left_index:base_right_index])
    right_data_str = ''
    if base_right_index < len(data_str):
        right_data_str = data_str[base_right_index:]
    n = 0
    while n < rp_num:
        m = 0
        while m < ic_num:
            new_data_str = left_data_str + str(base_num + m) + right_data_str
            yield new_data_str
            m += 1
        n += 1
def excel_writer(path):
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('sheet1')
    column1 = data_generate3('RT_三级1', 100, 10, 5, 6)
    column2 = data_generate2('87996650649', 1000, 0, 11)
    column3 = data_generate1('test', 1000)
    column4 = data_generate4('RT_三级1', 100, 10, 5, 6)
    i = 1
    for data1, data2, data3, data4 in zip(column1, column2, column3, column4):
        worksheet.write(i, 0, data1)
        worksheet.write(i, 1, data2)
        worksheet.write(i, 2, data3)
        worksheet.write(i, 3, data4)
        i += 1
    workbook.save(path)

if __name__ == '__main__':
    path = "C:/Users/ZXY/Desktop/test.xls"
    excel_writer(path)
