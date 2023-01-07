

from ctypes import windll, byref, create_unicode_buffer, create_string_buffer
FR_PRIVATE  = 0x10
FR_NOT_ENUM = 0x20


path=r'E:\Quiz App\Fonts\Ubuntu-Bold.ttf'
loadfont(path)


import tkinter
from  pyglet import  *
#
#
#  path='E:/Quiz App/Fonts/Ubuntu-Bold.ttf'
# font.add_file(path)
# ub=font.load("Ubuntu")
root = tkinter.Tk()
root.geometry("500x500")
MyLabel = tkinter.Label(root,text="Test",font=("Ubuntu Bold",25,"bold"))
MyLabel.pack()


MyLabel1 = tkinter.Label(root,text="test no fonts")
MyLabel1.pack()
root.mainloop()
