# coding=UTF-8
# -*-conding : gb2312 -*-


from tkinter import *
from tkinter.filedialog import askdirectory
import win32com.client
import win32api
import win32print
import time
import os

def selectPath():
    path_ = askdirectory()
    path.set(path_)

    #read file count
    file_name = os.listdir(path.get())
    file_dir = [os.path.join(path.get(), x) for x in file_name]
    count.set("file count = %d\n" % (len(file_dir)))


def startPrint():
    print("start print %s" %path.get())
    i = 0
    file_name = os.listdir(path.get())
    file_dirs = [os.path.join(path.get(), x) for x in file_name]

    while i < len(file_dirs):
        ext = os.path.splitext(file_dirs[i])[1]
        # 记录打印任务开始时间
        start_time = time.time()
        if ext.startswith('.x'):
            # excel
            xlApp = win32com.client.Dispatch('Excel.Application')
            xlApp.Visible = 0
            xlApp.EnableEvents = False
            xlApp.DisplayAlerts = False
            xlBook = xlApp.Workbooks.Open(file_dirs[i])
            xlApp.ActiveWorkbook.Sheets(1).PageSetup.Zoom = False
            xlApp.ActiveWorkbook.Sheets(1).PageSetup.FitToPagesWide = 1
            xlApp.ActiveWorkbook.Sheets(1).PageSetup.FitToPagesTall = 1
            xlBook.PrintOut(1, 99, )
            xlApp.quit()
        else:
            # word pdf txt
            win32api.ShellExecute(
                0,
                "print",
                file_dirs[i],
                '/d:"%s"' % win32print.GetDefaultPrinter(),
                ".",
                0
            )

        print(file_dirs[i])
        time.sleep(5)
        # 等待打印队列为空
        while True:
            default_printer = win32print.GetDefaultPrinter()
            printer_handle = win32print.OpenPrinter(default_printer)
            printer_info = win32print.GetPrinter(printer_handle, 2)  # 使用级别2
            win32print.ClosePrinter(printer_handle)

            # 直接访问字典来获取打印作业数量
            if printer_info['cJobs'] == 0:
                break
            time.sleep(5)  # 可以根据实际情况调整等待时间

        # 记录打印任务结束时间
        end_time = time.time()
        print(f"Printing took {end_time - start_time:.2f} seconds.")
        i = i + 1


root = Tk()

root.title('Print Tool')

width = 400
height = 280
screenwidth = root.winfo_screenwidth()
screenheight = root.winfo_screenheight()
alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
root.geometry(alignstr)
root.resizable(width=True, height=True)

count  = StringVar()



path = StringVar()
Label(root,text = "").pack()
Label(root,text = "Support batch printing of PDF, excel and word ",fg='red').pack(pady=10)
Label(root,text = "path:" ).pack()
Entry(root, textvariable = path).pack()
Button(root, text = "choose...", command = selectPath).pack(pady=10)
countLable = Label(root,textvariable = count,fg='blue').pack()
Button(root,text="start",command=startPrint).pack()

Label(root,text="author:zmy").pack()
Label(root,text="wechat:zmycloud").pack()

root.mainloop()
