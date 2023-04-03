import tkinter as tk
from tkinter import filedialog
root=tk.Tk()
root.withdraw()
Fpath=filedialog.askopenfilename()
#选一个文件获得路径


import xlwings as xw
app=xw.App(visible=True,add_book=False)
#不显示Excel消息框
app.display_alerts=False
#关闭屏幕更新,可加快宏的执行速度
app.screen_updating=False
# wk=xw.Book(Fpath)
# # 打开选的材料表
# muban=xw.Book('cailiaomuban.xls')
# # 打开模板
wk = app.books.open(Fpath)
# wk=xw.Book(Fpath)
# 打开选的材料表
# muban=xw.Book('cailiaomuban.xls')
muban = app.books.open('cailiaomuban.xls')
muban.save('材质清单.xls')
# 新建一个清单
f1=wk.sheets('封面')
# 读取材料表的封面
f2=muban.sheets('封面')
# 读取模板表的封面

biaoti1=f1.range('A2').value
# 复制标题
f2.range('A2').value=biaoti1
# 粘贴标题

ybiaoti1=f1.range('A3').value
# 复制英文标题
f2.range('A3').value=ybiaoti1
# 粘贴英文标题


area1=f1.range('A6').value
# 复制区域
f2.range('A6').value=area1
# 粘贴区域

area1=f1.range('A7').value
# 复制区域英文
f2.range('A7').value=area1
# 粘贴区域英文

data1=f1.range('A11').value
# 复制日期
f2.range('A11').value=data1
# 粘贴日期
#封面完成

qd1=wk.sheets('清单')
# 读取材料表的清单
qd2=muban.sheets('清单')
# 读取模板表的清单

info1=qd1.used_range
#获取原表信息
nrows1=info1.last_cell.row
#获取原表行数
ncols1=info1.last_cell.column
#获取原表列数
n=0
while n< (nrows1-20):
    qd2.api.Rows(7).Insert()
    n+=1

#插入行数,数量为原材料表-20




bh1= qd1.range((19,2),(nrows1,4)).options(ndim=2).value    # 读取二维的数据
qd2.range('a6').value = bh1                                  #复制二维数据

#复制材料编号,名称,区域

bh2= qd1.range((19,10),(nrows1,12)).options(ndim=2).value    # 读取二维的数据
qd2.range('e6').value = bh2                                   #复制二维数据

#复制型号,终饰,规格尺寸
bh3= qd1.range((19,5),(nrows1,5)).options(ndim=2).value    # 读取二维的数据
qd2.range('h6').value = bh3                                   #复制二维数据

#复制燃烧性能等级

bh4= qd1.range((19,13),(nrows1,15)).options(ndim=2).value    # 读取二维的数据
qd2.range('i6').value = bh4                                   #复制二维数据

#复制图案,色彩,描述

bh5= qd1.range((19,8),(nrows1,9)).options(ndim=2).value    # 读取二维的数据
qd2.range('l6').value = bh5                                   #复制二维数据

#复制供应商,生产商

bh6= qd1.range((19,18),(nrows1,18)).options(ndim=2).value    # 读取二维的数据
qd2.range('n6').value = bh6                                   #复制二维数据

#复制联系人

bh7= qd1.range((19,6),(nrows1,6)).options(ndim=2).value    # 读取二维的数据
qd2.range('r6').value = bh7                                   #复制二维数据

#复制最后修订日期

bh8= qd1.range((19,17),(nrows1,17)).options(ndim=2).value    # 读取二维的数据
qd2.range('q6').value = bh8                                  #复制二维数据

#复制最后修订日期

bh9= qd1.range((19,1),(nrows1,1)).options(ndim=2).value    # 读取二维的数据
qd2.range('s6').value = bh9                                  #复制二维数据

#复制标签种类


muban.save()
muban.close()
wk.close()
app.quit()
#全关掉

