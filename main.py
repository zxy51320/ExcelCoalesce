import tkinter as tk
import scr2 as scr
from tkinter import messagebox
from tkinter import filedialog

win =tk.Tk()
win.title('都撒到死啊')
win.geometry('720x450')
win.resizable(0,0)
#创建路径字典
path_dict = {'old_output_path':"",'input_path':"",'new_output_path':""}

def selectpath(indicator):
    # 从本地选择一个文件，并返回文件的目录
    _Path = filedialog.askdirectory()
    if _Path != '':
        if indicator == 1:
            lb1.config(text= _Path)
            path_dict['old_output_path'] = _Path #保存文件名及路径
        if indicator == 2:
            lb2.config(text= _Path)
            path_dict['input_path'] = _Path #保存文件名及路径
        if indicator == 3:
            lb3.config(text= _Path)
            path_dict['new_output_path'] = _Path #保存文件名及路径
            
def run():
    deals = getSelect()
    date = [date_input1.get(),date_input2.get()]
    if lb1['text']!='' and lb2['text']!='' and lb3['text']!='' and date!='' and len(deals)!=0:#参数验证
        if scr.openXLSX(path_dict,deals,date):
            messagebox.showinfo('','Job completed!')
        else:
            messagebox.showinfo('','Job failed! Please check logs on: ' + path_dict['new_output_path'])
    else:
        messagebox.showinfo('','Missing parameters!')

def getSelect():
    selected=[i.get() for i in v if i.get()]
    return selected


#deals复选框
list1=['GEMB','ECF','BNP 1','BNP 2','BNP 3','JET Money','BARLOG','MONETA I','MONETA II','BNP 4','MONETA III','JET Money II','ECF II','BNP 5','ECF III','MONETA IV','BNP 6']
v=[]
for index,item in enumerate(list1):
    v.append(tk.StringVar())
    tk.Checkbutton(win,text=item,variable=v[-1],onvalue=item,offvalue='').grid(row=index//5+1,column=index%5,sticky='w',padx=10)
    v[index].set(item)

#输入月份
lb0 = tk.Label(win,text='Last Month:')
lb0.grid(row=5,column=0,pady=10,padx=10)
default_date1 = tk.StringVar(value='2022-9')
date_input1 = tk.Entry(win,textvariable = default_date1, width=12)
date_input1.grid(row=5,column=1,pady=10,padx=10)
lb00 = tk.Label(win,text='This Month:')
lb00.grid(row=5,column=2,pady=10,padx=10)
default_date2 = tk.StringVar(value='2022-10')
date_input2 = tk.Entry(win,textvariable = default_date2, width=12)
date_input2.grid(row=5,column=3,pady=10,padx=10)
#路径标签
lb1 = tk.Label(win,text='')
lb1.grid(row=6,column=1,columnspan=10, pady=10, padx=10)
lb2 = tk.Label(win,text='')
lb2.grid(row=7,column=1,columnspan=10, pady=10, padx=10)
lb3 = tk.Label(win,text='')
lb3.grid(row=8,column=1,columnspan=10, pady=10, padx=10)
# 使用 grid()的函数来布局，并控制按钮的显示位置
old_sel_btn = tk.Button(win, text="Old Output Path", width=12, command=lambda: selectpath(1))
new_sel_btn = tk.Button(win, text="Input Path", width=12, command=lambda: selectpath(2))
out_sel_btn = tk.Button(win, text="New Output Path", width=12, command=lambda: selectpath(3))
old_sel_btn.grid(row=6, column=0, sticky="w", padx=10, pady=10)
new_sel_btn.grid(row=7, column=0, sticky="w", padx=10, pady=10)
out_sel_btn.grid(row=8, column=0, sticky="w", padx=10, pady=10)
run_btn = tk.Button(win, text="Run", width=12, command=run)
run_btn.grid(row=9, column=0, sticky="w", padx=10, pady=10)

win.mainloop()
