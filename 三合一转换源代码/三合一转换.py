from tkinter import filedialog
from 分程序.将病史txt转换成doc import 将病史txt转换成doc
from 分程序.将医嘱txt转换成excel import 将医嘱txt转换成excel
from 分程序.将检验txt转换成excel import 将检验txt转换成excel

path=filedialog.askopenfilenames(title='支持一次性选择多个文件：')
for i in path:
    if ('病史' in i)==True:
        将病史txt转换成doc(i)
    if ('医嘱' in i)==True:
        将医嘱txt转换成excel(i)
    if ('检验' in i)==True:
        将检验txt转换成excel(i)

input('按回车键退出')
