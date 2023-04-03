import tkinter as tk
from tkinter import filedialog
root=tk.Tk()
root.withdraw()
Fpath1=filedialog.askdirectory()
#选一个文件获夹得路径
print(Fpath1)
root=tk.Tk()
root.withdraw()
Fpath=filedialog.askopenfilename()
#选一个文件获得路径

import os
from PIL import Image
import xlwings as xw
app=xw.App(visible=True,add_book=False)
#不显示Excel消息框
app.display_alerts=False
#关闭屏幕更新,可加快宏的执行速度

wk=app.books.open(Fpath)

# 打开选的材料表

qd1=wk.sheets('内容')
# 读取材料表的清单
info1=qd1.used_range
nrows1=(info1.last_cell.row)/29


n=0
while n<nrows1:

    rng = qd1.range(f'A{15+n*29}')
    a=f'{Fpath1}/{n+1}.png'
    im=Image.open(a)
    _width,_height= im.size

    nw=520
    nh=520/_width*_height
    if nh>300:
        nh=300
        nw = 300 / _height*_width
    left = (525 - nw) / 2
    top = (rng.top) + (310 - nh) / 2

    n += 1
    print(os.path.abspath(a))
    a=os.path.abspath(a)
    qd1.pictures.add(a,left=left,top=top,height=nh)



wk.save()
wk.close()
app.quit()


#全关掉
# # left = rng.left + (rng.width - width) / 2 # 居中
# top = rng.top + (rng.height - height) / 2
