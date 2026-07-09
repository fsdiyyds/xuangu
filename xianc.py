# -*- coding: utf-8 -*-
"""
Created on Thu Nov  5 23:00:50 2020
8
123 124 125 121 119 122 126 123
@author: Administrator
"""

import _thread
import time
import tkinter as tk
from PIL import Image, ImageTk

window = tk.Tk()
window.geometry("900x600")
window.configure(bg ='pink')

img = Image.open('背景.jpg')
photo = ImageTk.PhotoImage(img)
tk.Label(window, image=photo).pack()



text = tk.Text(window,width=30,height=2)
text.pack()


def countdown(name1, a):
   b = a*60
   count = 0
   while count < b:
       text.insert(tk.INSERT, str(b-count))
       time.sleep(1)
       count += 1
       print (count)


def button1():
    _thread.start_new_thread(countdown,  ("Thread-1", 2, ))
        
def button2():
    _thread.start_new_thread(countdown,  ("Thread-2", 4, ))        



b1 = tk.Button(window,text = "Red",command = button1,activeforeground = "red",activebackground = "pink",pady=10)
b1.pack()


b2 = tk.Button(window,text = "Blue",command = button2,activeforeground = "red",activebackground = "pink",pady=10)
b2.pack()

# 进入消息循环
window.mainloop()

        
        
    
    
                    
        
    
    