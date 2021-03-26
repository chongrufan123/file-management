import easygui as g
import sys
from tkinter import *
import fatures as f
import xlrd

from tkinter import ttk

# 登录
class Login():
# 用户登录窗口,主要用easygui
    def login_interface(self):
        while 1:
            title = '校团委实践部文件管理系统(demo@Fan)'
            msg = '请选择你的身份'
            choice = ('管理员', '游客')
            identity = g.ccbox(msg, title, choice)  #管理员是TURE,游客是FALSE
            while(identity):
                msg = '请输入账号和密码'
                fields = ('账号', '密码')
                ret = g.multpasswordbox(msg, title, fields)
                if(ret == None):     # 点击右上角的xx
                    return 0
                if ret != ['admin', '123456']:   #现在暂时固定密码,之后可能会放到一个加密文件夹中
                    msg = '账号或密码输入错误,请选择重新输入还是返回上一步'
                    choice = g.ccbox(msg, title, choices = ('返回上一步', '重新输入'))
                    if(choice == None):
                        return 0
                    if(choice):
                        break
                    else:
                        continue
                else:
                    return 'admin'
            else:
                return 'tourist'

class Denglu():
    def __init__(self):
        self.count = 1
        self.a = f.Filen()
        self.get_data()
        self.delete_list = [[''],['']]  # 用来存放检索到的数据
        self.delete_list[0] = self.list[0]  #表头

        ## 从目标文件中获取数据存放到一个列表当中,之后加上如果目标文件不存在创建文件
    def  get_data(self):
        data = xlrd.open_workbook("../doc/shuju.xlsx")
        # 根据sheet名字来获取excel中的sheet
        table=data.sheets()[0]
        # 行数
        nrows=table.nrows
        # 某一行数据
        colnames=table.row_values(1)
        self.list=[]
        for rnum in range(0,nrows):
            self.list.append(table.row_values(rnum)[:13])  #只获取前13列,虽然不到13列
        
        ## 查看在文件中的所有数据,并展示出来
    def check(self):
        self.get_data()
        self.zhanshi(self.list)

        ## 将目标列表中的数据在UI上展示出来
    def zhanshi(self, listt):
        tt = Toplevel()
        tt.title("document")
        frame = Frame(tt, height=670, width=1845)

        tree = ttk.Treeview(frame, height=45)
        
        tree["columns"] = listt[0]
        for each in range(len(listt[0])):  #循环获得表头内容
            tree.column(listt[0][each], width=100)
            tree.heading(listt[0][each], text=listt[0][each])
                         
        for each1 in range(1, len(listt)):  #循环获得表格当中的内容
            tree.insert('', each1-1, text=each1-1, values=listt[each1])
        
        tree.pack()
        frame.pack()
        mainloop()
                
        # 调用增加一个元素的类
    def add_to1(self):
        self.a.add_to(self.nname.get())

        # 检索之后的确定按键,展示检索到的数据
    def determine(self):
        self.get_data()
        self.delete_list = [[''],['']]
        self.delete_list[0] = self.list[0]
#        print('文件名' + self.name.get())
#        print('年份' + self.vvar.get())
#        print('整理人'+self.persion.get())

       # 通过循环和多层条件判断确定是不是检索到的目标文件
       # 检索的方式是把列表中的数据强转为一个字符串去看字符串中是否包含这些文件,之后可能需要改进
        for each in range(1, len(self.list)):
            if self.name.get() in str(self.list[each]):
                
                if self.var.get() in str(self.list[each]):
                    if self.vvar.get() in str(self.list[each]):
                        if self.persion.get() in str(self.list[each]):
                            # 如果没有勾选主题的话会报错,所有这里把他放到最后,加个判断错误
                            try:
                                a = self.theme.get(self.theme.curselection()[0])
                                if a in str(self.list[each]):
                                    self.delete_list.append(self.list[each])
                            except:
                                self.delete_list.append(self.list[each])
                                
        self.rr.withdraw()   #退出这个子TK
#        print(self.delete_list)
        self.zhanshi(self.delete_list)
                                
    ## 检索文件的界面,之后可能将下面的语句和增加一个文件放到一起,他们很类似
    def retrieve(self):
        
        self.rr = Toplevel()
        self.rr.title("检索")
        
        frame1 = Frame(self.rr)
        frame2 = Frame(self.rr)

        w2 = Label(frame1, text = "文件检索，请选择文件的特征...", font = "30")
        w2.grid(row=0, column = 0, columnspan = 2, padx=5, pady=10, sticky=W)
        
        w2 = Label(frame1, text = "文件名", font = "30")
        w2.grid(row=1, column = 0, padx=5, pady=10, sticky=W)

        self.name = StringVar()
        nnamee = Entry(frame1, textvariable=self.name)
        nnamee.grid(row=1, column=1, padx=5,pady=10,stick=E)
        
        w2 = Label(frame1, text = "请选择文件的年份", font = "30")
        w2.grid(row=2, column = 0, padx=5, pady=10, sticky=W)
        self.var = StringVar()
        self.var.set("2021")
        testCMD = frame1.register(self.a.test1)
        year = Entry(frame1, textvariable=self.var, validate="key",\
                     validatecommand = (testCMD, '%P'))
        year.grid(row=2, column =1, padx=5, pady=10,sticky = E)

        w3 = Label(frame1, text = "实践主题", font = "30")
        w3.grid(row=3, column = 0, padx=5, pady=10, sticky=W)
        self.theme = Listbox(frame1, height = 5)
        self.theme.grid(row=3, column =1, padx=5, pady=10,sticky = E)
        with open("../doc/theme.txt") as f:
            for each in f:
                self.theme.insert(END, each.strip())


        options = []
        self.vvar = StringVar()
        
        with open("../doc/format.txt") as f:
            for each in f:
                options.append(each.strip())
                
        self.vvar.set(options[0])
        w4 = Label(frame1, text = "文件格式", font = "30")
        w4.grid(row=4, column = 0, padx=5, pady=10, sticky=W)
        format1 = OptionMenu(frame1, self.vvar, *options)
        format1.grid(row=4, column =1, padx=5, pady=10 ,sticky = E, ipadx = 10, ipady =0)
                

        w5 = Label(frame1, text = "整理人", font = "30")
        w5.grid(row=5, column = 0, padx=5, pady=10, sticky=W)
        self.persion = StringVar()
        note = Entry(frame1, textvariable=self.persion)
        note.grid(row=5, column =1, padx=5, pady=10,sticky = E)

        f = Button(frame2, text='确定', command = self.determine)
        f.grid(row=0, column=0, padx=5, pady=10, sticky=W)

        f = Button(frame2, text='返回', command = self.rr.withdraw)
        f.grid(row=0, column=1, padx=5, pady=10)

        f = Button(frame2, text='退出', command = self.rr.quit,\
                   background = 'red', foreground = 'white')
        f.grid(row=0, column=2, padx=5, pady=10, sticky=E)
        
        frame1.pack()
        frame2.pack()
    
        mainloop()

#    def delete(self):
        

        
        # 管理员之后的登录界面, 之后可能会加上删除数据的选项
    def admin_login(self):
        main = Tk()
            
        frame1 = Frame(main)
        frame2 = Frame(main)
        main.title('校团委实践部文件管理系统demo(@FAN)')

        #记录整理人的名字,之后的数据中会加上
        ww = Label(frame2, text = "请输入你的姓名", font = "30")
        ww.grid(row=0, column = 0, padx=5, pady=20, sticky=W)
        self.nname = StringVar()
        name = Entry(frame2, textvariable=self.nname)
    
        name.grid(row=0, column =1, padx=5, pady=10,sticky = E)
    
        w = Label(frame1, text="你好,管理员\n你可以选择以下选项", anchor=W, justify=LEFT, padx = 20, pady = 20, font= "50")
        w.grid(row = 0, sticky = W)

        b1 = Button(frame1, text='查询当前所有数据', command = self.check, padx = 10, pady = 10, font='50')
        b1.grid(row = 1, sticky = W, padx = 10, pady = 5)
        b2 = Button(frame1, text='添加新的数据', command = self.add_to1, padx = 10, pady = 10, font='50')
        b2.grid(row = 2, sticky = W, padx = 10, pady = 5)
#        b3 = Button(frame1, text='删除数据', command = self.delete, padx = 10, pady = 10, font='50')
 #       b3.grid(row = 3, sticky = W, padx = 10, pady = 5)
        b4 = Button(frame1, text='按条件检索数据', command = self.retrieve, padx = 10, pady = 10, font='50')
        b4.grid(row = 4, sticky = W, padx = 10, pady = 5)

        b4 = Button(frame1, text='退出', command = main.quit,\
                    background = 'red', foreground = 'white', font='50')
        b4.grid(row = 5, padx = 20, pady = 20)
        
        frame2.pack()
        frame1.pack()
    
        mainloop()

        # 游客登录界面,只有查看全部文件和检索文件两个选项,和上面的管理员登录界面的一部分内容可以放到一起
    def tourist_login(self):
        main = Tk()
            
        frame1 = Frame(main)
        frame2 = Frame(main)
        main.title('校团委实践部文件管理系统demo(@FAN)')
    
        w = Label(frame1, text="你好,游客\n你可以选择以下选项", anchor=W, justify=LEFT, padx = 20, pady = 20, font= "50")
        w.grid(row = 0, sticky = W)

        b1 = Button(frame1, text='查询当前所有数据', command = self.check, padx = 10, pady = 10, font='50')
        b1.grid(row = 1, sticky = W, padx = 10, pady = 5)
        
        b4 = Button(frame1, text='按条件检索数据', command = self.retrieve, padx = 10, pady = 10, font='50')
        b4.grid(row = 4, sticky = W, padx = 10, pady = 5)

        b4 = Button(frame1, text='退出', command = main.quit,\
                    background = 'red', foreground = 'white', font='50')
        b4.grid(row = 5, padx = 20, pady = 20)
        
        frame2.pack()
        frame1.pack()
    
        mainloop()
    

if  __name__ == '__main__':
    a = Denglu()
    a.tourist_login()
        

        
