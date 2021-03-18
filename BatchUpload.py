import tkinter as tk
import getpass
from tkinter import filedialog
from datetime import datetime
import threading
import os
import pymysql
# import socket
import xlrd
import time
import subprocess
import tkinter.messagebox as message_box

class SEGThread(threading.Thread):
    def __init__(self, func, *args):
        super().__init__()
        self.func = func
        self.args = args
        self.setDaemon(True)
        self.start()

    def run(self):
        self.func(*self.args)

def get_local_date():
    dt = datetime.now()
    return dt.strftime('%Y-%m-%d %H:%M:%S')

class Upload_GUI:
    def __init__(self):
        self.root = tk.Tk()
        # self.root.minsize(200,200)
        self.root.title("Batch upload")
        self.root.resizable(False, False)
        sw = self.root.winfo_screenwidth()  # get screenwidth
        sh = self.root.winfo_screenheight()  # get screenheight
        ww = 600  # set application window weidth
        wh = 400  # set application window height
        x = (sw - ww) / 2
        y = (sh - wh) / 2
        self.root.geometry("%dx%d+%d+%d" % (ww, wh, x, y))  # set window size

        self.frame1 = tk.LabelFrame(self.root,text="File List choose",width = 200)
        self.button1 = tk.Button(self.frame1, text="Browse", command=self.__file_browse)
        self.button1.pack(side=tk.LEFT)
        self.string = tk.StringVar()
        self.text = tk.Entry(self.frame1, width = 100, state = 'readonly', textvariable = self.string)
        self.text.pack(side=tk.RIGHT)
        self.frame1.pack()
        self.frame2 = tk.LabelFrame(self.root,text="Server Connect and File upload")
        self.button2 = tk.Button(self.frame2, text = 'Connect to Server',state = 'disabled', command=lambda: SEGThread(self.__connector_db))
        self.button2.pack(side=tk.LEFT)
        self.button3 = tk.Button(self.frame2, text="Upload",state = 'disabled', command=lambda: SEGThread(self.__upload))
        self.button3.pack(side=tk.RIGHT)
        self.frame2.pack()
        # self.frame = tk.Frame(self.root).pack()
        # self.canvas = tk.Canvas(self.frame,width = 600, height = 30, bg = 'white')
        # self.canvas.pack()
        # self.x = tk.StringVar()
        # self.out_rec = self.canvas.create_rectangle(5,5,550,25,outline = "blue",width = 1)
        # self.fill_rec = self.canvas.create_rectangle(5,5,5,25,outline = "",width = 0,fill = "blue")
        # tk.Label(self.frame,textvariable = self.x).pack
        self.frame3 = tk.LabelFrame(self.root,text="Process")
        self.scroll = tk.Scrollbar()
        self.scroll.pack(side = tk.RIGHT,fill = tk.Y)
        self.text1 = tk.Text(self.frame3)
        self.text1.pack()
        self.scroll.config(command=self.text1.yview)
        self.text1.config(yscrollcommand=self.scroll.set, state='disable')
        self.frame3.pack(side='bottom')
        self.excel_file = ''
        self.MySQL = None
        self.cursor = None
        self.excel_name = ''

    def __file_browse(self):

        self.text1.config(state='normal')
        self.text1.delete(0.0, tk.END)
        self.text1.config(state='disabled')
        user_name = getpass.getuser()
        initial_path = r'C:\Users\%s\Desktop\EN_Lib_Code\2.0 Project\Batch upload' % (user_name)
        if not os.path.exists(initial_path):
            initial_path = r'C:\Users\%s\Desktop' % (user_name)
        # initial_path = r'C:\Users\%s\Desktop' % (user_name)
        file_path = filedialog.askopenfilename(title='choose file list excel',
                                               filetypes=[('Excel', '*.xlsx'), ('All Files', '*')],
                                               initialdir=initial_path)
        # 在光标处插入文字
        if file_path != '':
            self.button2.config(state='normal')
            # self.button3.config(state='normal')
            self.text.config(state = 'normal')
            self.string.set(file_path)
            self.text.config(state='readonly')
            self.excel_file = r'%s' %(file_path)
            file_name = os.path.basename(file_path)
            self.excel_name = file_name
            string_xlsx = 'List file: '+ file_name + ' prepared'
            self.__insert_text(string_xlsx)
            data = xlrd.open_workbook(self.excel_file)
            table = data.sheets()[0]
            nor = table.nrows
            string_file = 'A total of ' + str(nor-1) + ' files need to be uploaded'
            self.__insert_text(string_file)
        else:
            self.button2.config(state='disabled')
            self.string.set(file_path)
            string = 'Please choose an valid file'
            self.__insert_text(string)
        self.button3.config(state='disabled')

    def __connector_db(self):
        string = 'Connecting server...'
        self.__insert_text(string)
        time.sleep(0.5)
        try:
            # 1.连接到mysql数据库
            # HostName = "SGHZ001013372"
            # IpAddress = socket.gethostbyname(HostName)
            # localhost连接本地数据库 user 用户名 password 密码 db数据库名称 charset 数据库编码格式
            # self.MySQL = pymysql.connect(host=IpAddress, user='fengbang', password='FBA2CS@2019', db='dblibrary',
            #                             charset='utf8')
            self.MySQL = pymysql.connect(host='10.219.129.77', user='root', password='HAOxue008', db='en_library',
                                         charset='utf8')
            Connected = True
            self.button3.config(state='normal')
        except:
            Connected = False
            self.MySQL = None

        if Connected:
            string = 'Connecting sucess'
        else:
            string = 'Connecting Failed, Please reconnect'
        self.__insert_text(string)



    def __upload(self):
        Excel_data = self.file_read()
        file_table_head = self.get_head_of_file_table()
        download = 0
        visitsnum = 0
        state = 3
        sendemail = 1
        table = 'file_table'
        i = 1
        logfile = self.create_logfile()
        # print(logfile)
        self.button1.config(state = 'disabled')
        self.button2.config(state = 'disabled')
        self.button3.config(state = 'disabled')

        for item in Excel_data:
            self.cursor = self.MySQL.cursor()
            string_err = ''

            date = get_local_date()
            path = item['Link'].strip()
            title = item['Title'].strip()
            category = item['Category'].strip()
            classification = item['Classification'].strip()
            keywords = item['Keywords'].strip()
            author = item['Author'].strip()
            deliver = item['Deliver'].strip()
            auditor = item['Auditor'].strip()
            customer = item['Customer'].strip()
            source = item['Source'].strip()
            participants = item['Participants'].strip()
            format = item['Format'].strip()
            # new_file_path = r'%s\%s.%s' % (path, title, format)
            compare = self.check_same_title(path, title, format, category)
            # print(compare)
            if compare == 1:
                string = str(i) + ': ' + title + ' already in Database'
                self.__insert_text(string)
                self.write_log(logfile,string)
            else:
                try:
                    # create No of file
                    file_number, version = self.create_file_number(category, classification,title)

                    # insert file into file_table
                    values = [file_number, category, classification, version, title, date, author, deliver, \
                              participants, source, customer, format, keywords, visitsnum, auditor, state, sendemail]
                    self.insert_data(table, file_table_head, values)
                    time.sleep(1)
                    # transfer file into Library
                    self.file_transefer(path, title, format, category, file_number)
                    string = str(i) + ': ' +title + ' sucessfully uploaded. ' + 'File number: '+ file_number
                    # self.__insert_text(string)
                    # self.write_log(logfile, string)
                except:
                    string = str(i) + ': ' + title + ' transfer failed.' + ' Please check your excel file'
                finally:
                    self.__insert_text(string)
                    self.write_log(logfile, string)

            i = i+1
            self.text1.see(tk.END)
            self.cursor.close()
        self.MySQL.close()
        string = 'File upload finshed' + ' You can click to browse start an new upload.'
        self.__insert_text(string)
        self.button1.config(state = 'normal')
        sentence = 'Upload files from ' + self.excel_name + ' finished'
        self.show_message_box(sentence)
        # self.button2.config(state = 'normal')
        # self.button3.config(state = 'normal')
    def show_message_box(self,sentence):
        top = tk.Tk()
        top.withdraw()
        top.update()
        message_box.showinfo('Info', sentence)

    def check_same_title(self,path, title, format, category):
        sql = "select * from file_table where Title = '%s' and Format = '%s'" %(title, format)
        cursor = self.MySQL.cursor()
        if cursor.execute(sql):
            f1 = r'%s\%s.%s' % (path, title, format)
            f1size = self.get_fileSize(f1)
            all = cursor.fetchall()
            f2_no = all[-1][0]
            f2_title = all[-1][4]
            f2_format = all[-1][11]
            sql = "select Path from category_table where Category = '%s'" %(category)
            cursor.execute(sql)
            f2_path = cursor.fetchall()[0][0]
            f2 = r'%s\%s_%s.%s' % (f2_path, f2_no, f2_title, f2_format)
            f2size = self.get_fileSize(f2)
            if f1size == f2size:
                compare = 1
            else:
                compare = 0
        else:
            compare = 0
        cursor.close()
        return compare

    def get_fileSize(self,f):
        fsize = os.path.getsize(f)
        fsize = fsize / float(1024 * 1024)

        return round(fsize, 2)

    def file_read(self):
        data = xlrd.open_workbook(self.excel_file)
        table = data.sheets()[0]
        nor = table.nrows
        # print(nor)
        self.total = nor-1
        nol = table.ncols
        dict = {}
        for i in range(1, nor):
            for j in range(nol):
                title = table.cell_value(0, j)
                value = table.cell_value(i, j)
                dict[title] = value
            yield dict
        return dict

    def get_head_of_file_table(self):
        sql = 'select * from file_table'
        self.cursor = self.MySQL.cursor()
        self.cursor.execute(sql)
        self.cursor.close()
        fields = self.cursor.description
        head = []
        # 或取数据库中表头
        for field in fields:
            head.append(field[0])
        return head

    def create_file_number(self,category, classification, title):
        # confirm category code
        sql = "select Code from category_table where Category = '%s'" % (category)
        # print(sql)
        self.cursor.execute(sql)
        CategoryCode = self.cursor.fetchall()[0][0]
        # print(CategoryCode)

        # confirm calssification code
        sql = "select Code from classification_table where Category = '%s' and Classification = '%s'" % (
            category, classification)
        self.cursor.execute(sql)
        CalssificationCode = self.cursor.fetchall()[0][0]
        # print(CalssificationCode)

        # confirm version
        sql = "select Version from file_table where state=3 and Title = '%s'" % (title)
        try:
            self.cursor.execute(sql)
            version = self.cursor.fetchall()[0][0]
            version = str(int(version) + 1)
        except:
            version = '01'
        if len(version) == 1:
            version = '0' + version
        elif len(version) > 2:
            print('Version is beyond 99, please rename the title!')
        # print(version)

        # confirm file number
        sql = 'select No from file_table where Category = "%s" and Classification = "%s" and State = 3 ' \
                  'order by Date desc limit 1' % (category, classification)
        try:
            self.cursor.execute(sql)
            number = self.cursor.fetchall()[-1][0][-10:]
            new_num = str(int(number) + 1)
            # print(new_num)
        except:
            new_num = '1'
        # print(new_num)
        Number_of_file = CategoryCode + CalssificationCode + version + (10 - len(new_num)) * '0' + new_num
        # print(Number_of_file)
        return Number_of_file, version

    def insert_data(self,table, file_table_head, values):
        sql = 'insert into %s(' % table + \
              ','.join([i for i in file_table_head]) + ') values (' + ','.join(repr(i) for i in values) + ')'
        try:
            self.cursor.execute(sql)
            self.MySQL.commit()
        except:
            self.MySQL.rollback()

    def file_transefer(self, path, title, format, category, file_number):
        old_file = r'%s\%s.%s' % (path, title, format)
        new_file_path = self.get_new_file_path(category)
        new_file_name = r'%s\%s_%s.%s' % (new_file_path,
                                          file_number,
                                          title,
                                          format)
        cmd = 'copy "%s" "%s"' % (old_file, new_file_name)
        subprocess.call(cmd, shell=True, timeout=10)

    def get_new_file_path(self,category):
        sql = "select path from category_table where Category = '%s'" % (category)
        self.cursor.execute(sql)
        path = self.cursor.fetchall()[0][0]
        return path

    def __insert_text(self,string):
        self.text1.config(state='normal')
        insert_text = string + '\n'
        self.text1.insert("insert", insert_text)
        self.text1.config(state='disable')
        string = ''

    def write_log(self, logfile, string):
        # insert = [title, file_number]
        # print(string)
        fp = open(logfile,'a',encoding='utf-8')
        fp.writelines(string+'\r\n')
        # fp.write()
        fp.close()

    def create_logfile(self):
        file_name = r'.\\'+ os.path.basename(self.excel_file[:-4])+'upload.txt'
        if os.path.exists(file_name):
            os.remove(file_name)
        return file_name



if __name__ == "__main__":
	GUI = Upload_GUI()
	GUI.root.mainloop()
