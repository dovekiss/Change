#打开清单 打开模板 数下需要做几个标签页 创建标签页并命名 数下每个标签页的材料数量 创建表格 输入每个材料的编号 删除模板 保存
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
wk=xw.Book(Fpath)
# 打开选的材料表
muban=xw.Book('jiajuwuliaoshumuban.xls')
# 打开模板
muban.save('家具技术规格书.xls')
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
print(nrows1)
n=0
while n< (nrows1-7):
    qd2.api.Rows(20).Insert()
    n+=1

#插入行数,数量为原清单-7




bh1= qd1.range((6,1),(nrows1-1,3)).options(ndim=2).value    # 读取二维的数据
qd2.range('b19').value = bh1                                  #复制二维数据

#复制材料编号,名称,区域

bh2= qd1.range((6,5),(nrows1-1,7)).options(ndim=2).value    # 读取二维的数据
qd2.range('j19').value = bh2                                   #复制二维数据

#复制型号,终饰,规格尺寸
bh3= qd1.range((6,8),(nrows1-1,8)).options(ndim=2).value    # 读取二维的数据
qd2.range('s19').value = bh3                                   #复制二维数据

#复制燃烧性能等级

bh4= qd1.range((6,9),(nrows1-1,11)).options(ndim=2).value    # 读取二维的数据
qd2.range('m19').value = bh4                                   #复制二维数据

#复制图案,色彩,描述

bh5= qd1.range((6,12),(nrows1-1,13)).options(ndim=2).value    # 读取二维的数据
qd2.range('h19').value = bh5                                   #复制二维数据

#复制供应商,生产商

bh6= qd1.range((6,14),(nrows1-1,14)).options(ndim=2).value    # 读取二维的数据
qd2.range('r19').value = bh6                                   #复制二维数据

#复制联系人

bh7= qd1.range((6,18),(nrows1-1,18)).options(ndim=2).value    # 读取二维的数据
qd2.range('f19').value = bh7                                   #复制二维数据

#复制最后修订日期

bh8= qd1.range((6,17),(nrows1-1,17)).options(ndim=2).value    # 读取二维的数据
qd2.range('q19').value = bh8                                  #复制二维数据

#复制修订次数

# bh9= qd1.range((6,19),(nrows1-1,19)).options(ndim=2).value    # 读取二维的数据
# qd2.range('s19').value = bh9                                  #复制二维数据
#
# #复制标签种类

bh10= qd1.range((6,16),(nrows1-1,16)).options(ndim=2).value    # 读取二维的数据
qd2.range('p19').value = bh10                                  #复制二维数据

#复制制作日期种
n=0
while n< (nrows1-6):
    qd2.range(f'g{19+n}').value=f'第{n+1}页'
    n+=1



nr1=muban.sheets('内容')
# 读取模板表的内容页
n=0
while n< ((nrows1-7)*29):
    nr1.api.Rows(30).Insert()
    n+=1
n=0
while n< (nrows1-7):
    nr1.api.Rows(f'{1+28*n+n}:{29+28*n+n}').Copy(nr1.api.Rows(f'{30+28*n+n}'))
    n+=1
n=0
while n< (nrows1-6):
    nr1.range(f'H{28+28*n+n}').value=qd2.range(f'B{19+n}').value
    n+=1

info2=nr1.used_range
#获取原表信息
nrows2=info2.last_cell.row+1
#获取原表行数
nr1.api.Rows(nrows2).Delete()
muban.save()
muban.close()
wk.close()
app.quit()
#全关掉