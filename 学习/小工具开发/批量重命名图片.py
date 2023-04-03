import os
import tkinter as tk
from tkinter import filedialog
root=tk.Tk()
root.withdraw()
Fpath1=filedialog.askdirectory()
#选一个文件获夹得路径

filelist = os.listdir(Fpath1) #该文件夹下所有的文件（包括文件夹） 作者：杭漂一族小张 https://www.bilibili.com/read/cv15156613/ 出处：bilibili
a = 0
for file in filelist:
    Olddir = os.path.join(Fpath1, file)

    filename=os.path.splitext(file)[0] #分离文件名与扩展名;得到文件名
    filename = filename[3:]

    Newdir = os.path.join(Fpath1, str(filename) + str('.png'))  # 得到路径+后面需要跟的命名 作者：杭漂一族小张 https://www.bilibili.com/read/cv15156613/ 出处：bilibili
    a += 1
    os.rename(Olddir,Newdir)  # 重命名 作者：杭漂一族小张 https://www.bilibili.com/read/cv15156613/ 出处：bilibili