import os
import xlrd

config_file_name = "config.xlsx"  # 配置文件名
workbook = xlrd.open_workbook(config_file_name)
sheet = workbook.sheet_by_index(0)  # 获取sheet，索引从0开始
cols = sheet.col_values(0)  # 拿到关键字名单

listFileName = cols[0]  # 花名册文件名
dirName = cols[1]  # 作业文件夹名
file_name = cols[2]  # 输出文件名
col = int(cols[3])  # 关键字列，0开始数
questNameCol = int(cols[4])  # 输出关键字列，0开始

if os.path.exists(file_name):
    os.remove(file_name)

# 拿名单
workbook = xlrd.open_workbook(listFileName)
sheet = workbook.sheet_by_index(0)  # 获取sheet，索引从0开始
cols = sheet.col_values(col)  # 拿到比对关键字名单
questNames = sheet.col_values(questNameCol)  # 拿到输出名单

test_names = os.listdir(dirName)  # 拿作业名单

with open(file_name, "w") as file:
    index = 0
    for i in cols:  # 名单名
        if any(i in test_name for test_name in test_names):
            index += 1
        else:
            file.write(str(questNames[index]))
            file.write('\n')
            file.flush()
            index += 1

file.close()
print("打印成功,请查看生成文件")
input()
