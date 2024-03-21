import time
import xlwt
import xlrd
import webbrowser
import os


# f = open(c + '.bat', 'a+')
# # # f.write(('\n' + a))
# f.close()
# 不创建#r只读 r+可读写
# 创建，覆盖#w只写w+可读写
# 创建，不覆盖#a只写a+可读可写

# 创建workbook（其实就是excel，后来保存一下就行）
# workbook = xlwt.Workbook(encoding='ascii')
workbook = xlwt.Workbook(encoding='UTF-8')
# 创建表
worksheet = workbook.add_sheet('My Worksheet')
# 往单元格内写入内容
x = 0
y = 0
a = 1
renshu = 56

# worksheet.write(0, 0, label='写入')

# 初始化
# 初始化
# 初始化
def chushihua():
    # 初始化
    print('初始化中————')
    worksheet.write(0, 0, label='学号')
    worksheet.write(0, 1, label='姓名')
    worksheet.write(0, 2, label='学科名称')
    f = open('a.txt', 'r')
    # f.write(('\n' + a))
    for i in range(renshu):
        worksheet.write(i+1, 0, label=i+1)
        worksheet.write(i+1,1,label=f.readline(i+4))
    f.close()

# 随机输入
# 随机输入
# 随机输入
def suiji():
    print('输入‘-1’退出')
    # print('学号+成绩')
    while True:
        a = input('学号：')
        if (a == '-1'):
            break
        else:
            b = input('成绩：')
            # if (b == '-1'):
            #     break
            # else:
            worksheet.write(int(a), 2, label=b)
    # 保存
    workbook.save('文件.xls')

print('‘完全随机’初始化完成！')


# 顺序输入
# 顺序输入
# 顺序输入
def moshi1():
    print('输入‘-1’退出')
    # print('学号+成绩')
    a = 0
    while True:
        a = (a + 1)
        print(('学号：' + str(a)))
        b = input('成绩：')
        if (b == '-1'):
            break
        else:
            worksheet.write(int(a), 2, label=b)
    # 保存
    workbook.save('文件.xls')

print('‘模式1’初始化完成！')


def suiji():
    a = 1
    print('输入‘-1’退出')
    # print('学号+成绩')
    for i in range(renshu):
        import random
        print('学号：' + str(a))
        b = random.randint(0, 100)
        print('成绩：' + str(b))
        worksheet.write(int(a), 2, label=b)
        a = a + 1
    # 保存
    workbook.save('文件.xls')

print('‘随机’初始化完成！')


# 打开
# 打开
# 打开
def dakai():
    print(('文件路径' + os.path.realpath('文件.xls')))
    print('保存成功，是否打开？(T/F)')
    if (input('') == 'T'):
        webbrowser.open('file://'+os.path.realpath('文件.xls'))

print('‘打开’初始化完成！')

while True:
    time.sleep(0.2)
    print('请选择：')
    time.sleep(0.2)
    print('1.随机输入')
    time.sleep(0.2)
    print('2.顺序输入')
    time.sleep(0.2)
    print('3.详细介绍')
    time.sleep(0.2)
    print('4.完全随机输入')
    time.sleep(0.2)
    print('5.关闭')
    time.sleep(0.2)
    xuanze = input('请输入：')
    if (xuanze == '1'):
        chushihua()
        moren()
        dakai()
    elif (xuanze == '2'):
        chushihua()
        moshi1()
        dakai()
    elif (xuanze == '3'):
        print('随机输入：学号和成绩都可以自定义')
        time.sleep(0.2)
        print('顺序输入：只有成绩都可以自定义，学号按1一直往后排')
        time.sleep(0.2)
        print('完全随机输入：学号、成绩完全随机')
        time.sleep(0.2)
        print('关闭：关闭当前程序（Ctrl+C再输入y也可以关闭）')
        time.sleep(0.2)
        input('看完随便输入字符：')
    elif (xuanze == '4'):
        chushihua()
        suiji()
        dakai()
    elif (xuanze == '5'):
        break
    else:
        print('φ(゜▽゜*)♪，你好像输入错了，再输入一次吧！（输入数字）')
        time.sleep(1.5)
