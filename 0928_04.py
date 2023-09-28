'''
# This program is written by 李泽钧 & 彭瑞安
# For the purpose of dualing with Excel(merge and deduplicate excel for more specific) in China Telecom work procedure
#
# Date: 2023/9/28 
'''


import pandas as pd
import time
import os

start_time = time.time()
print("\n去重小程序\n作者：李泽钧 彭瑞安\n")
print("\n=====================================\n")

excel_path = input("请输入Excel文件路径：")
# 读取Excel文件
print("\n正在读取Excel文件...\n")
read_starttime = time.time()

df = pd.read_excel(excel_path, sheet_name='基站承载视图1')

read_endtime = time.time()
read_time = read_endtime-read_starttime
print(f"读入耗时:{read_time:.2f}秒\n")
# 编辑某一列
#df['column2'] = df['column2'].apply(lambda x: x.split('.')[0] if '.' in x else x)
df['设备端口'] = df['设备端口'].astype(str).apply(
    lambda x: x.split('.')[0] if '.' in x else x)
print("\n正在去除Vlan后缀...\n")
# 将两列合并为一列
df['new_column'] = df['设备名称'] + df['设备端口']

# 仅保留第一行
df = df.drop_duplicates(subset='new_column', keep='first')
print("\n正在去重，请稍后..\n")

# 将结果写回Excel文件

# 获取用户主目录  
home_directory = os.path.expanduser("~")  
  
# 提供一个输入选项让他们选择在哪个位置保存Excel文件  
save_location = input("请输入保存Excel文件的完整路径（例如：/home/user/Documents/），或者直接按Enter键在默认位置（{})下保存：".format(home_directory))  
  
# 如果用户没有输入任何内容，那么就使用默认位置  
if not save_location:  
    save_location = home_directory  
  
# 确保路径存在，如果不存在则创建  
if not os.path.exists(save_location):  
    os.makedirs(save_location)

#df.to_excel('your_result_file.xlsx', index=False)
df.to_excel(os.path.join(save_location, "your_result_file.xlsx"))
finish_time = time.time()
total_time = finish_time-start_time
print(f"总耗时:{total_time:.2f}秒\n")
print("\n去重完成，10秒后将关闭程序。\n")
time.sleep(10)
input()
