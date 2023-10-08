import pandas as pd  
import tkinter as tk  
from tkinter import filedialog  
from tkinter import messagebox  
import time  
import os
  
class App:  
    def __init__(self, master):  
        self.master = master  
        master.title("去重小程序")  
  
        self.text_box1 = tk.Text(master)  
        self.text_box1.pack()  
  
        self.button1 = tk.Button(master, text="选择文件", command=self.select_file)  
        self.button1.pack()  
  
        self.button2 = tk.Button(master, text="选择输出文件夹", command=self.select_folder)  
        self.button2.pack()  
  
        self.text_box2 = tk.Text(master)  
        self.text_box2.pack()  
  
        self.start_time = None  
        self.end_time = None  
        self.total_time = None  
  
    def select_file(self):  
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])  
        if file_path:  
            self.text_box1.delete('10.0', tk.END)  # 清空文本框  
            self.text_box1.insert(tk.END, "文件路径： " + file_path + "\n")  # 在文本框输出文件路径  
            self.start_time = time.time()  # 记录开始时间  
            self.process_data(file_path)  # 处理数据...  
  
    def select_folder(self):  
        folder_path = filedialog.askdirectory()  
        if folder_path:  
            self.text_box2.delete('10.0', tk.END)  # 清空文本框  
            self.text_box2.insert(tk.END, "输出文件夹路径： " + folder_path + "\n")  # 在文本框输出文件夹路径  
  
    def process_data(self, file_path, folder_path):
        read_starttime = time.time()
        df = pd.read_excel(file_path, sheet_name='基站承载视图1')
        read_endtime = time.time()
        read_time = read_endtime-read_starttime  
        self.text_box1.insert(tk.END, f"读取用时： {read_time:.2f}秒\n正在处理中...\n")  # 在文本框中显示读取用时和“正在处理中”字样
        '''
        df = pd.read_excel(file_path)  # 读取Excel文件  
        df = df[df["基站承载视图1"].str.contains(".") == False]  # 过滤掉包含"."的行  
        df = df.dropna(how='all')  # 删除空值行  
        df = df[df["基站承载视图1"].duplicated(keep='first')]  # 删除重复行，只保留第一行  
        '''
        start_time = time.time()  # 记录开始时间
        df['设备端口'] = df['设备端口'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
        # 将两列合并为一列
        df['new_column'] = df['设备名称'] + df['设备端口']
        # 仅保留第一行
        df = df.drop_duplicates(subset='new_column', keep='first')  
        #output_file = self.text_box2.get('10.0', tk.END) + "/" + file_path[file_path.rfind("/")+1:].replace(".xlsx", "") + ".csv"  # 输出文件路径  
        #df.to_csv(output_file, index=False)  # 将处理后的表格输出到指定文件夹中  
        df.to_excel(os.path.join(folder_path, "your_result_file.xlsx"))
        end_time = time.time()  # 记录结束时间  
        total_time = end_time - start_time  # 计算总耗时  
          
        self.end_time = time.time()  # 记录结束时间  
        self.total_time = total_time  # 计算总耗时  
        messagebox.showinfo("处理完成", "处理完成，总耗时： {:.2f}秒".format(total_time))  # 在消息框中显示总耗时  
  
if __name__ == '__main__':  
    root = tk.Tk()  
    # 创建一个Label，显示作者信息  
    author_label = tk.Label(root, text="作者：李泽钧 彭瑞安", font=("Arial", 6), anchor="w")  
    author_label.pack(fill=tk.X, ipadx=5, pady=5)  # 使用padx和pady设置Label与其他组件的间距
    app = App(root)  
    root.mainloop()