from tkinter import *
from tkinter import filedialog
import os
import xlwt
import xlrd
from xlutils.copy import copy

class Filen():
    def __init__(self):
        self.filename = '文件'

    def callback(self):
        fileame = filedialog.askopenfilename()
        self.filename = fileame
        self.v.set(self.filename)

    def test1(self, content):
        return content.isdigit()

    def yes(self):
        if(not os.path.isfile("../doc/shuju.xlsx")):
           book = xlwt.Workbook()
           sheet = book.add_sheet("汇总文件表")
           sheet.write(0,0,"文件名")
           sheet.write(0,1,"路径")
           sheet.write(0,2,"年份")
           sheet.write(0,3,"文件主题")
           sheet.write(0,4,"文件格式")
           sheet.write(0,5,"备注")
           sheet.write(0,6,"整理人")
           book.save("../doc/shuju.xlsx")
        word_book = xlrd.open_workbook("../doc/shuju.xlsx")
        sheets = word_book.sheet_names()
        work_sheet = word_book.sheet_by_name(sheets[0])
        old_rows = work_sheet.nrows
        new_work_book = copy(word_book)
        new_sheet = new_work_book.get_sheet(0)
        i = old_rows
        for each in self.list1:
            for j in range(len(self.list1)):
                new_sheet.write(i, j, self.list1[j])
        new_work_book.save('../doc/shuju.xlsx')
        self.top1.withdraw()

    def determine(self):
        self.top1 = Toplevel()
        self.top1.title("请确认")
        top = Frame(self.top1)
        top2 = Frame(self.top1)
        filell = os.path.basename(self.filename)     #去除路径，返回文件名称
        filellen = "文件名称: " + filell
        la1 = Label(top, text = filellen, font ="30")
        la1.grid(row=0, column=0, padx=5, pady=10, sticky=W)

        filesen = "文件路径: " + self.filename
        la2 = Label(top, text = filesen, font ="30")
        la2.grid(row=1, column=0, padx=5, pady=10, sticky=W)
        
        yearsen = "年份: " + self.var.get()
        la3 = Label(top, text = yearsen, font ="30")
        la3.grid(row=2, column=0, padx=5, pady=10, sticky=W)
        
        themesen = "文件主题: " + self.theme.get(self.theme.curselection()[0])
        la4 = Label(top, text = themesen, font ="30")
        la4.grid(row=3, column=0, padx=5, pady=10, sticky=W)

        format1sen = "文件格式: " + self.vvar.get().strip()
        la5 = Label(top, text = format1sen, font ="30")
        la5.grid(row=4, column=0, padx=5, pady=10, sticky=W)

        beizhusen = "备注: " + self.beizhu.get().strip()
        la6 = Label(top, text = beizhusen, font ="30")
        la6.grid(row=5, column=0, padx=5, pady=10, sticky=W)

        writersen = "整理人: " + self.writer
        la7 = Label(top, text = writersen, font ="30")
        la7.grid(row=6, column=0, padx=5, pady=10, sticky=W)

        f = Button(top2, text='确定', command = self.yes)
        f.grid(row=0, column=0, padx=5, pady=10, sticky=W)

        f = Button(top2, text='返回', command = self.top1.withdraw)
        f.grid(row=0, column=1, padx=5, pady=10, sticky =E)

        self.list1 = [filell, self.filename, self.var.get(), self.theme.get(self.theme.curselection()[0]), self.vvar.get().strip(), self.beizhu.get(), self.writer]
        top.pack()
        top2.pack()

        
    def return_list(self):
        pass
        
    def add_to(self, wirter):
        self.writer = wirter
        master = Toplevel()
        master.title("校团委实践部文件管理系统demo(@FAN)")
        frame1 = Frame(master)
        frame2 = Frame(master)
        frame3 = Frame(master)
        w = Label(frame1, text = "请选择要添加的文件或文件夹", font ="30")
        w.grid(row=0, column=0, padx=5, pady=10, sticky=W)
        f = Button(frame1, text='打开文件', command = self.callback)
        f.grid(row=0, column=1, padx=5, pady=10, sticky=E)
        print(self.filename)
        self.v = StringVar()
    
        self.v.set(self.filename)
        finame = Label(frame1, textvariable=self.v)
        finame.grid(row=1, column=0,columnspan=2, padx=5, pady=10, sticky=W)

        w2 = Label(frame1, text = "请选择文件的年份", font = "30")
        w2.grid(row=3, column = 0, padx=5, pady=10, sticky=W)
        self.var = StringVar()
        self.var.set("2021")
        testCMD = frame1.register(self.test1)
        year = Entry(frame1, textvariable=self.var, validate="key",\
                     validatecommand = (testCMD, '%P'))
        year.grid(row=3, column =1, padx=5, pady=10,sticky = E)

        w3 = Label(frame1, text = "实践主题", font = "30")
        w3.grid(row=4, column = 0, padx=5, pady=10, sticky=W)
        self.theme = Listbox(frame1, height = 5)
        self.theme.grid(row=4, column =1, padx=5, pady=10,sticky = E)
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
        w4.grid(row=5, column = 0, padx=5, pady=10, sticky=W)
        format1 = OptionMenu(frame1, self.vvar, *options)
        format1.grid(row=5, column =1, padx=5, pady=10 ,sticky = E, ipadx = 10, ipady =0)
                

        w5 = Label(frame1, text = "备注", font = "30")
        w5.grid(row=6, column = 0, padx=5, pady=10, sticky=W)
        self.beizhu = StringVar()
        note = Entry(frame1, textvariable=self.beizhu)
        note.grid(row=6, column =1, padx=5, pady=10,sticky = E)

        f = Button(frame2, text='确定', command = self.determine)
        f.grid(row=0, column=0, padx=5, pady=10, sticky=W)

        f = Button(frame2, text='返回', command = master.withdraw)
        f.grid(row=0, column=1, padx=5, pady=10)

        f = Button(frame2, text='退出', command = master.quit,\
                   background = 'red', foreground = 'white')
        f.grid(row=0, column=2, padx=5, pady=10, sticky=E)
        
        frame1.pack()
        frame2.pack()
        frame3.pack()
    
    
        mainloop()

def delete():
    pass

def retrieve():
    pass

if __name__ == "__main__":
    a = Filen()
    a.add_to()
