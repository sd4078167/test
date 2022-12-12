import os;import pandas as pd;import pandas;import xlrd;import tkinter as tk;from tkinter import filedialog;import tkinter.messagebox
#新建一个询问窗口,提示怎么用和选择功能
windows1 = tk.Tk()
windows1.withdraw
tk.Label(windows1,text='"-----使用说明-----"  \n 分别输入1~2来实现功能 \n 1.合并表格 \n 2.拆分表格').pack()
tk.Label(windows1,text='请输入相应的数字：').pack()
值 = tk.Entry(windows1)
值.pack()
def 输入的值():
    global 获取值
    获取值 = 值.get()
    获取值 = int(获取值)
    windows1.destroy()
tk.Button(windows1,text='确认',command=输入的值).pack()
windows1.mainloop()

# 合并的函数
def 合并():
    root = tk.Tk()
    root.withdraw
    tk.messagebox.showwarning('提示', '请选择需要汇总的文件')
    文件集合 = tk.filedialog.askopenfilenames()
    print(文件集合)
    wb = pandas.DataFrame()
    k = tk.Entry(root)
    k.pack()
    起始行 = k
    # input('输入标题在第几行:')
    for file in 文件集合:
        工作表名 = pandas.ExcelFile(file)           #pandas.ExcelFile 和read_excel的区别，是仅读取，READ_EXCEL是读取成DATAFRAME的数据结构
        print(工作表名)
        for sheet in 工作表名.sheet_names:
            nwb = pandas.read_excel(file,sheet_name=sheet, header=int(起始行) - 1)
            wb = pandas.concat([wb,nwb])
    print(wb)
    tk.messagebox.showinfo('提示', '请选择你想保存的位置')
    保存路径 = tk.filedialog.askdirectory()
    wb.to_excel(保存路径 + '/' + '合并结果.xlsx',index=False)

# def 合并工作簿():
#     tk.messagebox.showwarning('提示', '请选择需要汇总的文件')
#     文件集合 = tk.filedialog.askopenfilenames()
#     print(文件集合)
#     wb = pandas.DataFrame()
#     起始行 = input('输入标题在第几行:')
#     for file in 文件集合:
#         nwb = pandas.read_excel(file,header=int(起始行) - 1)
#         wb = pandas.concat([wb,nwb])
#     print(wb)
#     tk.messagebox.showinfo('提示', '请选择你想保存的位置')
#     保存路径 = tk.filedialog.askdirectory()
#     wb.to_excel(保存路径 + '/' + '结果.xlsx',index=False)

# def 合并工作表():
#     wb = pandas.DataFrame()
#     tk.messagebox.showinfo('提示', '请选择要汇总的工作表所在的工作簿')
#     工作簿路径 = filedialog.askopenfilename()
#     起始行 = input('输入标题在第几行:')
#     工作表名 = pandas.ExcelFile(工作簿路径)
#     for sheet in 工作表名.sheet_names:
#         nwb = pandas.read_excel(工作簿路径,sheet_name=sheet, header=int(起始行) - 1)
#         wb = pandas.concat([wb,nwb])
#     print(wb)
#     tk.messagebox.showinfo('提示', '请选择你想保存的位置')
#     保存路径 = tk.filedialog.askdirectory()
#     print(保存路径)
#     wb.to_excel(保存路径 + '/' + '结果.xlsx',index=False)

#拆分的函数
def 拆分():
    windows = tk.Tk()
    windows.withdraw
    表名 = tk.filedialog.askopenfilename()
    tk.messagebox.showwarning('填写说明', '请填写需要拆分表的标题所在的行,以及拆分的参数')
    tk.Label(windows, text='标题在第几行').pack()
    标题行数1 = tk.Entry(windows)
    标题行数1.pack()
    tk.Label(windows, text='参数1(必须是标题行的某一标题)').pack()
    分割参数3 = tk.Entry(windows)
    分割参数3.pack()
    tk.Label(windows, text='参数2(必须是标题行的某一标题)').pack()
    分割参数4 = tk.Entry(windows)
    分割参数4.pack()
    def close():
        global 标题行数;
        global 分割参数;
        global 分割参数2
        标题行数 = 标题行数1.get()
        分割参数 = 分割参数3.get()
        分割参数2 = 分割参数4.get()
        windows.destroy()
    按钮 = tk.Button(windows, text='确定并关闭窗口', command=close)
    按钮.pack()
    windows.mainloop()
    tk.messagebox.showinfo('提示', '请选择保存在什么位置')
    保存路径 = tk.filedialog.askdirectory()
    os.chdir(保存路径+'/')
    file = pd.read_excel(表名, header=int(标题行数) - 1)
    grad = list(file[分割参数].drop_duplicates())
    for i in grad:
        date = file[file[分割参数] == i]
        biao = pd.ExcelWriter(i + '.xlsx')
        Class = list(date[分割参数2].drop_duplicates())
        for C in Class:
            date1 = date[date[分割参数2] == C]
            date1.to_excel(biao, sheet_name=str(C), index=False)
        biao.save()
        biao.close()
if 获取值 == 1:
    合并()
elif 获取值 == 2:
    拆分()