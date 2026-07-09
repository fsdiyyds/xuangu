# -*- coding: utf-8 -*-

'''
@Time    : 2022/2/13 19:42
@Author  : fenglei
@FileName: nianhui.py
@Software: PyCharm
 
'''

import _thread
import random
import re
import time
import tkinter as tk  # 使用Tkinter前需要先导入
import winsound
from PIL import Image, ImageTk

window = tk.Tk()  # 第1步，实例化object，建立窗口window
# 第2步，给窗口的可视化起名字
window.title('你来比划我来猜     -------桥二所2023年春节活动')
# 第3步，设定窗口的大小(长 * 宽)
window.geometry('1240x1000')  # 这里的乘是小x
window.configure(bg="Tomato")
# 第4步，在图形界面上设定标签
l = tk.Label(window, text='请比划该词语：', bg='Tomato', foreground="yellow", font=('黑体', 30), width=130,
             height=2).place(x=-500, y=1)
words = open("words.txt", 'r', encoding='UTF-8').read()
words = re.split(r'[\n 、]', words)

img = Image.open('背景2.jpg')
photo = ImageTk.PhotoImage(img)
tk.Label(window, image=photo, anchor="center").place(x=20, y=565)


def selectukuang():
    print(words)
    count = len(words)
    id = random.randint(0, count - 1)
    print(id)
    word = str(words[id])
    words.pop(id)
    l1 = tk.Label(window, text=word, font=('黑体', 96), width=15,
                  height=2, background="SkyBlue").place(x=230, y=250)
    return word


def countdown(name, min):
    winsound.PlaySound(r"合并.wav", winsound.SND_FILENAME | winsound.SND_ASYNC)
    count = 0
    b = min * 60
    l2 = tk.Text(window, width=7, height=1, background="pink", foreground="green")
    l2.place(x=1150, y=200)
    l2.config(font=("黑体", 30, "bold"))

    while count < b:
        ncount = b - count
        # text.pack(fill=tkinter.X, side=tkinter.BOTTOM)
        l2.delete(1.0, tk.END)
        l2.insert(tk.END, str(ncount))  # INSERT表示在光标位置插入
        l2.see(tk.END)
        l2.update()
        time.sleep(1)
        if ncount < 60:
            winsound.Beep(1800, 100)
        count += 1
    l2.delete(1.0, tk.END)
    l2.insert(tk.END, "时间到！")  # INSERT表示在光标位置插入
    winsound.Beep(35000, 1500)
    l2.see(tk.END)
    l2.update()


def jishi():
    _thread.start_new_thread(countdown, ("Thread-1", 2,))


button_tukuang = tk.Button(window, text="下一词", command=selectukuang, activeforeground="red",
                           activebackground="pink").place(x=700, y=200)

button_jishi = tk.Button(window, text="计时开始", command=jishi, activeforeground="blue", activebackground="green").place(
    x=1000, y=200)

window.mainloop()
