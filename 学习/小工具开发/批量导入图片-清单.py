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
app.screen_updating=False
wk=app.books.open(Fpath)

# 打开选的材料表

qd1=wk.sheets('清单')
# 读取材料表的清单
info1=qd1.used_range
nrows1=info1.last_cell.row




n=0
while n<(nrows1-6):
    n += 1
    rng = qd1.range(f'D{n+5}')
    a=f'{Fpath1}/{n}.png'
    im=Image.open(a)
    _width,_height= im.size

    nw=120
    nh=120/_width*_height
    if nh>90:
        nh=90
        nw = 90 / _height*_width
    left = rng.left + (rng.width - nw) / 2
    top = rng.top + (rng.height - nh) / 2

    a = os.path.abspath(a)
    qd1.pictures.add(a,left=left,top=top,height=nh)



wk.save()
wk.close()
app.quit()


#全关掉
# # left = rng.left + (rng.width - width) / 2 # 居中
# top = rng.top + (rng.height - height) / 2
