# %%
import xlwings
from difflib import get_close_matches as gcm
from os import mkdir
import numpy as np
from fuzzywuzzy import fuzz as fz
import traceback


#这是一个完整的python程序的主文件，用于调用其他文件中的脚本，实现整个程序的功能
#其它两个文件分别为task1.py和task2.py
#这个文件中的函数是用于实现程序的调用功能
def main():

    print('Hello World!')
    print('This is a python program.')
    print('It is used to call other python scripts.')
    print('输入1，调用task1.py')
    print('输入2，调用task2.py')
    print('输入3，退出程序')
    selc = 0
    while True:
        selc = input('请输入：')
        while selc not in ['1', '2', '3']:
            selc = input('输入错误，请重新输入：')
            # selc = input('请输入：')

        if selc == '3':
            break
        
        with open(f'task{selc}.py', 'r',encoding='utf-8') as f:
            code = f.read()
            try:
                exec(code, globals()) # 这里的globals()是为了让task.py中的函数可以访问到task.py中的变量
                # 捕获异常，防止程序出错

            except:
                traceback.print_exc()
                print(f'task{selc}.py','执行失败!\n')
            f.close()
            print('执行结束')
    print('程序结束')
    print()

    pass
#这里是调用其他文件中的函数，实现程的功能

if __name__ == '__main__':
    main()