import pandas as pd  
import tkinter as tk  
from tkinter import filedialog  
from tkinter import messagebox  
import time  
import os  
  
class Application(tk.Frame):  
    def __init__(self, master=None):
        master.title("去重过滤小程序")  
        super().__init__(master)  
        self.master = master  
        self.pack()  
        self.create_widgets()  
  
    def create_widgets(self):
        '''  
        self.start_button = tk.Button(self)  
        self.start_button["text"] = "选择文件和文件夹"  
        self.start_button["command"] = self.ask_for_file_folder  
        self.start_button.pack(side="top")  '''
  
        self.input_button = tk.Button(self)  
        self.input_button["text"] = "选择输入文件"  
        self.input_button["command"] = self.ask_for_input_file  
        self.input_button.pack(side="top")  
  
        self.output_button = tk.Button(self)  
        self.output_button["text"] = "选择输出文件夹"  
        self.output_button["command"] = self.ask_for_output_folder  
        self.output_button.pack(side="top")  
  
        self.start_button2 = tk.Button(self)  
        self.start_button2["text"] = "开始"  
        self.start_button2["command"] = self.process_data  
        self.start_button2.pack(side="top")  
  
        self.text_box = tk.Text(self)  
        self.text_box.pack(side="top", fill="both", expand=True)  
  
        self.text_box2 = tk.Text(self)  
        self.text_box2.pack(side="bottom", fill="both", expand=True)  
    '''
    def ask_for_file_folder(self):  
        filetypes = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]  
        filename = filedialog.askopenfilename(filetypes=filetypes)  
        if filename:  
            self.input_file = filename  
            self.output_folder = filedialog.askdirectory()  
            if self.input_file and self.output_folder:  
                self.start_button2["state"] = "normal"  
            else:  
                messagebox.showinfo("错误", "请确保您已选择输入文件和输出文件夹")  
    '''
    def ask_for_input_file(self):  
        filetypes = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]  
        filename = filedialog.askopenfilename(filetypes=filetypes)  
        if filename:  
            self.input_file = filename
            self.text_box2.insert(tk.END, "需要处理的Excel文件路径： " + filename + "\n")  # 在文本框输出文件路径  
            if self.output_folder:  
                self.start_button2["state"] = "normal"  
            else:  
                messagebox.showinfo("错误", "请确保您已选择输出文件夹")  
  
    def ask_for_output_folder(self):  
        foldername = filedialog.askdirectory()  
        if foldername:  
            self.output_folder = foldername
            self.text_box2.insert(tk.END, "输出文件夹路径： " + foldername + "\n")  # 在文本框输出文件夹路径  
            if self.input_file:  
                self.start_button2["state"] = "normal"  
            else:  
                messagebox.showinfo("错误", "请确保您已选择输入文件")  
  
    def process_data(self):  
        start_time = time.time()  
        messagebox.showinfo("提示", "正在处理中...")  
        df = pd.read_excel(self.input_file, sheet_name='基站承载视图1')  
        duration = time.time() - start_time  
        self.text_box.insert(tk.END, f"读取用时： {duration:.2f}秒\n")  
        self.text_box.insert(tk.END, "正在处理中...\n")  
        #start_time = time.time()

        # 删除"所有者"列中含有特定字符串的所在行  
        #df = df[~df['基站厂家'].str.contains('联通')]
        #df= df[(~df['基站厂家'].str.contains('联通', na=False)) | (df['基站厂家'].isna())] 
        #df = df[~df['基站厂家'].str.contains('联通', na=False)] 
  
        # 删除"状态"列为"down"的所在行  
        #df = df[df['端口状态'] != 'down']   
        df['设备端口'] = df['设备端口'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
        # 将两列合并为一列
        df['new_column'] = df['设备名称'] + df['设备端口']
        # 仅保留第一行
        df = df.drop_duplicates(subset='new_column', keep='first')
        df = df[~df['基站厂家'].str.contains('联通', na=False)] #删除“基站厂家”列含有联通的所在行
        # 删除"状态"列为"down"的所在行
        df = df[df['端口状态'] != 'down']  
        df.to_excel(os.path.join(self.output_folder, "output_" + time.strftime("%Y%m%d%H%M") + ".xlsx"))
        duration = time.time() - start_time  
        self.text_box.insert(tk.END, f"总耗时： {duration:.2f}秒\n")  
        self.text_box.insert(tk.END, "处理完成\n")

if __name__ == '__main__':  
    root = tk.Tk()  
    # 创建一个Label，显示作者信息  
    author_label = tk.Label(root, text="作者：李泽钧 彭瑞安", font=("Arial", 8), anchor="w")  
    author_label.pack(fill=tk.X, ipadx=5, pady=5)  # 使用padx和pady设置Label与其他组件的间距
    author_label2 = tk.Label(root, text="请确保Excel表中含有名为“基站承载视图1”的sheet", font=("Arial", 12), anchor="center")  
    author_label2.pack(fill=tk.X, ipadx=5, pady=5)  # 使用padx和pady设置Label与其他组件的间距
    author_label3 = tk.Label(root, text="此程序是先进行去重，再进行过滤含有联通字符或状态为down的行,去重和过滤顺序不同会导致结果不同", font=("Arial", 10), anchor="center")  
    author_label3.pack(fill=tk.X, ipadx=5, pady=5)  # 使用padx和pady设置Label与其他组件的间距
    app = Application(root)  
    root.mainloop()