##quản lý thư viện v1.2
#modules
import customtkinter as ctk
import tkinter,webbrowser
from tkinter import *
from tkinter import messagebox,ttk
from datetime import date
import os,pathlib,openpyxl,xlrd
from openpyxl import Workbook
import requests,sys
from bs4 import BeautifulSoup
import openpyxl
from pathlib import Path
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
ds_admin=[]
my_file = Path(".file_need\\admin.txt")
if not my_file.is_file():
    my_file.touch()
with open('.file_need\\admin.txt','w') as n:
    n.write('Admin')
    n.close()
with open('.file_need\\admin.txt','r') as f:
    f=f.read()
    a=f.split(',')
    for i in range(len(a)):
        ds_admin.append(a[i])
file=pathlib.Path('DANH_SACH_MUON.xlsx')
if file.exists():
    pass
else:
    file=Workbook()
    sheet=file.active
    sheet['A1']='Đăng ký số:'
    sheet['B1']='Tên'
    sheet['C1']='Lớp'
    sheet['D1']='Giới Tính'
    sheet['E1']='Ngày Đăng Ký'
    sheet['F1']='Địa chỉ'
    sheet['G1']='Thông Tin Liên Lạc'
    sheet['H1']='Tên Sách Mươn'
    sheet['I1']='Mượn Ngày:'
    sheet['J1']='Ngày Trả'
    sheet['K1']='Trách Nhiệm'
    sheet['L1']='Người Xét Duyệt'
    file.save('DANH_SACH_MUON.xlsx')
class CTKAPP(ctk.CTk):
	def __init__(self, *args, **kwargs):
		ctk.CTk.__init__(self, *args, **kwargs)
		container = ctk.CTkFrame(self)
		container.pack(side = "top", fill = "both", expand = True)
		container.grid_rowconfigure(0, weight = 1)
		container.grid_columnconfigure(0, weight = 1)
		self.frames = {}
		for F in (StartPage, Page1, Page2,ADMIN):
			frame = F(container, self)
			self.frames[F] = frame
			frame.grid(row = 0, column = 0, sticky ="nsew")
		self.show_frame(StartPage)
	def show_frame(self, cont):
		frame = self.frames[cont]
		frame.tkraise()
class StartPage(ctk.CTkFrame):
	def __init__(self, parent, controller):
		ctk.CTkFrame.__init__(self, parent)
		frame1=ctk.CTkFrame(self,width=450,height=30)
		frame1.pack_configure(side='bottom')
		bt1=ctk.CTkButton(frame1,text='Home',command=lambda:messagebox.showinfo('Information','Bạn Đang ở Home'))
		bt1.place(x=0,y=0)
		bt2=ctk.CTkButton(frame1,text='Danh sách Người Mượn',command=lambda:controller.show_frame(Page1))
		bt2.place(x=150,y=0)
		bt3=ctk.CTkButton(frame1,text='Cài Đặt',command=lambda:controller.show_frame(Page2))
		bt3.place(x=310,y=0)
class Page1(ctk.CTkFrame):
	def __init__(self, parent, controller):
		ctk.CTkFrame.__init__(self, parent)
		frame1=ctk.CTkFrame(self,width=450,height=30)
		frame1.pack_configure(side='bottom')
		bt1=ctk.CTkButton(frame1,text='Home',command=lambda:controller.show_frame(StartPage))
		bt1.place(x=0,y=0)
		bt2=ctk.CTkButton(frame1,text='Danh sách Người Mượn',command=lambda:messagebox.showinfo('Information','Bạn Đang ở trang này'))
		bt2.place(x=150,y=0)
		bt3=ctk.CTkButton(frame1,text='Cài Đặt',command=lambda:controller.show_frame(Page2))
		bt3.place(x=310,y=0)
		tree=ctk.CTkScrollbar(self)
		tree.pack()
		trees=ctk.
class Page2(ctk.CTkFrame):
	def __init__(self, parent, controller):
		ctk.CTkFrame.__init__(self, parent)
		ctk.CTkLabel(self,text='Cài Đặt').pack_configure(side='top')
		frame1=ctk.CTkFrame(self,width=450,height=30)
		frame1.pack_configure(side='bottom')
		frame2=ctk.CTkFrame(self)
		frame2.pack_configure(side='left')
		bt1=ctk.CTkButton(frame1,text='Home',command=lambda:controller.show_frame(StartPage))
		bt1.place(x=0,y=0)
		bt2=ctk.CTkButton(frame1,text='Danh sách Người Mượn',command=lambda:controller.show_frame(Page1))
		bt2.place(x=150,y=0)
		bt3=ctk.CTkButton(frame1,text='Cài Đặt',command=lambda:messagebox.showinfo('Information','Bạn Đang ở trang này'))
		bt3.place(x=310,y=0)
		def reset_all():
			msg=messagebox.askquestion('Infomation','Khôi phục cài đặt gốc sẽ xoá toàn bộ dữ liệu kể cả dữ liệu    người mượn.Xác nhận xoá?')
			if msg=='yes':
				os.remove('DANH_SACH_MUON.xlsx')
				os.remove('.file_need//admin.txt')
		frame_bt1=ctk.CTkButton(frame2,text='Khôi Phục Cài Đặt Gốc',command=reset_all)
		frame_bt1.pack()
		frame_bt2=ctk.CTkButton(frame2,text='Quản lý Admin',command=lambda:controller.show_frame(ADMIN))
		frame_bt2.pack(pady=15)
		frame_bt3=ctk.CTkButton(frame2,text='Kiểm tra cập nhật')
		frame_bt3.pack(pady=10)
class ADMIN(ctk.CTkFrame):
	def __init__(self, parent, controller):
		ctk.CTkFrame.__init__(self, parent)
		frame1=ctk.CTkFrame(self,width=450,height=30)
		frame1.pack_configure(side='bottom')
		lb1=ctk.CTkLabel(self,text='Quản Lý Admin')
		lb1.pack()
		bt1=ctk.CTkButton(frame1,text='Home',command=lambda:controller.show_frame(StartPage))
		bt1.place(x=0,y=0)
		bt2=ctk.CTkButton(frame1,text='Danh sách Người Mượn',command=lambda:controller.show_frame(Page1))
		bt2.place(x=150,y=0)
		bt3=ctk.CTkButton(frame1,text='Cài Đặt',command=lambda:controller.show_frame(Page2))
		bt3.place(x=310,y=0)

app = CTKAPP()
app.title('Quản lý thư viện')
app.iconbitmap('.ico//logo.ico')
app.mainloop()