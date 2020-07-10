from tkinter import *
from tkinter import ttk
import pyodbc
import csv
import re
import pathlib
from datetime import datetime

server = "DESKTOP-9LPUIR9"
databaseName = "QUAN_LY_CONTAINER"

driver = '{SQL Server}' # Driver you need to connect to the database
port = '1433'
connection = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server}; SERVER='+server+';DATABASE='+databaseName+';Trusted_Connection=yes;')

cursor = connection.cursor()

# print(pathlib.Path().absolute())

# Đổ data vào TABLE Container
def showdataTab1():
    cursor.execute("""SELECT SO_CONTAINER,
                    (
                     SELECT CM.NAME 
                     FROM COM_CODE CM
                     WHERE CN.LOAI_CONTAINER = CM.CODE
                    ) LOAI_CONTAINER,
                    (
                     SELECT CM.NAME 
                     FROM COM_CODE CM
                     WHERE CN.TINH_TRANG = CM.CODE
                    ) TINH_TRANG,
                    SUBSTRING(NGAY_VAO, 1,4) + '-' + SUBSTRING(NGAY_VAO, 5,2) + '-' + SUBSTRING(NGAY_VAO, 7,2) AS NGAY_VAO_CANG,
                    (
                     SELECT CM.NAME 
                     FROM COM_CODE CM
                     WHERE CN.XUAT_NHAP = CM.CODE
                    ) XUAT_NHAP_KHAU,
                    CONG_TY,
                    NGAY_TAO
                    
                    FROM CONTAINER CN
                    WHERE DELETE_FLAG = 'N'""")
    tempList=cursor.fetchall()
    for row in listBox1.get_children():
        listBox1.delete(row)
    for col1, col2, col3, col4, col5, col6, col7 in tempList:
        listBox1.insert("", "end", values=(col1, col2, col3, col4, col5, col6, col7))

# Đổ data lên TABLE Booking
def showdataTab2():
    cursor.execute("""SELECT SO_BOOKING, 
                   SO_CONTAINER,
                   SUBSTRING(NGAY_VAO, 1,4) + '-' + SUBSTRING(NGAY_VAO, 5,2) + '-' + SUBSTRING(NGAY_VAO, 7,2) AS NGAY_VAO_CANG,
                   CONG_TY,
                   NGAY_TAO
                   FROM BOOKING""")

    tempList=cursor.fetchall()
    for row in listBox2.get_children():
        listBox2.delete(row)
    for col1,col2,col3,col4,col5 in tempList:
        listBox2.insert("", "end", values=(col1, col2, col3, col4, col5))


# Row selection
def fillEntryTab1(event):
    for i in listBox1.selection():
        list = listBox1.item(i, 'values')
        soContainer1.delete(0,END)
        loaiContainer1.delete(0,END)
        status1.delete(0,END)
        namVaoCang1.delete(0,END)
        thangVaoCang1.delete(0,END)
        ngayVaoCang1.delete(0,END)
        xuatNhapKhau1.delete(0,END)
        congTy1.delete(0,END)

        soContainer1.insert(END, list[0])
        loaiContainer1.insert(END, list[1])
        status1.insert(END, list[2])
        xuatNhapKhau1.insert(END, list[4])
        congTy1.insert(END,list[5])

        col4 = list[3].split("-")
        namVaoCang1.set(col4[0])
        thangVaoCang1.set(col4[1])
        ngayVaoCang1.set(col4[2])


def fillEntryTab2(event):
    for i in listBox2.selection():
        list = listBox2.item(i, 'values')
        soBooking2.delete(0,END)
        soContainer2.delete(0,END)
        namVaoCang2.delete(0,END)
        thangVaoCang2.delete(0,END)
        ngayVaoCang2.delete(0,END)
        congTy2.delete(0,END)

        soBooking2.insert(END, list[0])
        soContainer2.insert(END, list[1])
        congTy2.insert(END, list[3])

        col3 = list[2].split("-")
        namVaoCang2.set(col3[0])
        thangVaoCang2.set(col3[1])
        ngayVaoCang2.set(col3[2])

def clearTextTab1():
        soContainer1.delete(0,END)
        loaiContainer1.delete(0,END)
        status1.delete(0,END)
        namVaoCang1.delete(0,END)
        thangVaoCang1.delete(0,END)
        ngayVaoCang1.delete(0,END)
        xuatNhapKhau1.delete(0,END)
        congTy1.delete(0,END)

def clearTextTab2():
        soBooking2.delete(0,END)
        soContainer2.delete(0,END)
        namVaoCang2.delete(0,END)
        thangVaoCang2.delete(0,END)
        ngayVaoCang2.delete(0,END)
        congTy2.delete(0,END)

def insertIntoTab1():
    cursor.execute("""INSERT INTO CONTAINER VALUES('{0}'
                   ,(SELECT CODE FROM COM_CODE WHERE NAME = N'{1}')
                   ,(SELECT CODE FROM COM_CODE WHERE NAME = N'{2}')
                   ,(SELECT '{3}'+'{4}'+'{5}')
                   ,(SELECT CODE FROM COM_CODE WHERE NAME = N'{6}')
                   ,'{7}'
                   ,'N'
                   , GETDATE())
                   """.format(soContainer1.get(), loaiContainer1.get(), status1.get(),
                              namVaoCang1.get(),thangVaoCang1.get().zfill(2),ngayVaoCang1.get().zfill(2),
                              xuatNhapKhau1.get(), congTy1.get()))
    connection.commit()
    showdataTab1()


def insertIntoTab2():
    if is_valid_container_id(soContainer2.get()):
        cursor.execute("""INSERT INTO BOOKING VALUES('{0}','{1}'
                   ,(SELECT '{2}'+'{3}'+'{4}')
                   ,'{5}'
                   , GETDATE())
                   """.format(soBooking2.get(), soContainer2.get(),
                              namVaoCang2.get(),thangVaoCang2.get().zfill(2),ngayVaoCang2.get().zfill(2),
                              congTy2.get()))
        connection.commit()
        showdataTab2()
        clearTextTab2()
        labelWarning.grid_remove()
    else:
        labelWarning.grid()


def updateTab1():
    cursor.execute(
        """
        UPDATE CONTAINER
        SET LOAI_CONTAINER = (SELECT CODE FROM COM_CODE WHERE NAME =N'{1}')
        , TINH_TRANG = (SELECT CODE FROM COM_CODE WHERE NAME = N'{2}')
        , NGAY_VAO = (SELECT '{3}'+'{4}'+'{5}')
        , XUAT_NHAP = (SELECT CODE FROM COM_CODE WHERE NAME = N'{6}')
        , CONG_TY = '{7}'
        , NGAY_TAO = GETDATE()
        WHERE SO_CONTAINER = '{0}'
        """.format(soContainer1.get(), loaiContainer1.get(), status1.get(),
                   namVaoCang1.get(), thangVaoCang1.get().zfill(2), ngayVaoCang1.get().zfill(2),
                   xuatNhapKhau1.get(), congTy1.get())
        )
    connection.commit()
    showdataTab1()
    clearTextTab1()


def updateTab2():
    if is_valid_container_id(soContainer2.get()):
        cursor.execute(
            """
            UPDATE BOOKING
            SET SO_CONTAINER = '{1}'
            , NGAY_VAO = (SELECT '{2}'+'{3}'+'{4}')
            , CONG_TY = '{5}'
            , NGAY_TAO = GETDATE()
            WHERE SO_BOOKING = '{0}'
            """.format(soBooking2.get(), soContainer2.get(), namVaoCang2.get(), thangVaoCang2.get().zfill(2), ngayVaoCang2.get().zfill(2), congTy2.get())
        )
        connection.commit()
        showdataTab2()
        clearTextTab2()
        labelWarning.grid_remove()
    else:
        labelWarning.grid()


def deleteFlagTab1():
    cursor.execute(
        """
        UPDATE CONTAINER
        SET DELETE_FLAG ='Y'
        WHERE SO_CONTAINER ='{0}'
        """.format(soContainer1.get()))
    connection.commit()
    showdataTab1()
    clearTextTab1()

def deleteRowTab2():
    cursor.execute(""" DELETE FROM BOOKING WHERE SO_BOOKING = '{0}' """.format(soBooking2.get()))
    connection.commit()
    showdataTab2()
    clearTextTab2()


def exportExcelTab1():
    try:
        cursor.execute("""SELECT SO_CONTAINER,
                    (
                     SELECT CM.NAME 
                     FROM COM_CODE CM
                     WHERE CN.LOAI_CONTAINER = CM.CODE
                    ) LOAI_CONTAINER,
                    (
                     SELECT CM.NAME 
                     FROM COM_CODE CM
                     WHERE CN.TINH_TRANG = CM.CODE
                    ) TINH_TRANG,
                    SUBSTRING(NGAY_VAO, 1,4) + '-' + SUBSTRING(NGAY_VAO, 5,2) + '-' + SUBSTRING(NGAY_VAO, 7,2) AS NGAY_VAO_CANG,
                    (
                     SELECT CM.NAME 
                     FROM COM_CODE CM
                     WHERE CN.XUAT_NHAP = CM.CODE
                    ) XUAT_NHAP_KHAU,
                    CONG_TY,
                    NGAY_TAO
                    
                    FROM CONTAINER CN
                    WHERE DELETE_FLAG = 'N'""")

        now = datetime.now().timestamp()
        with open('Container' + str(now) + '.csv', 'w', encoding='utf-8', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)
        connection.commit()

    except:
        lblExcel.grid()


def exportExcelTab2():
    try:
        cursor.execute("""SELECT SO_BOOKING, 
                   SO_CONTAINER,
                   SUBSTRING(NGAY_VAO, 1,4) + '-' + SUBSTRING(NGAY_VAO, 5,2) + '-' + SUBSTRING(NGAY_VAO, 7,2) AS NGAY_VAO_CANG,
                   CONG_TY,
                   NGAY_TAO
                   FROM BOOKING""")

        now = datetime.now().timestamp()
        with open('Booking' + str(now) + '.csv', 'w', encoding='utf-8', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow([i[0] for i in cursor.description])
            csv_writer.writerows(cursor)
        connection.commit()
    except:
        lblExcel.grid()


def is_valid_container_id(container_id):
    pattern = '[a-zA-Z]{4}[0-9]{7}'
    result = re.search(pattern, container_id)
    return result is not None


def checkBtn():
    if is_valid_container_id(soContainer2.get()):
        labelWarning.grid_remove()
    else:
        labelWarning.grid()


# CREATE FORM
newForm = Tk()
newForm.geometry('900x500')
newForm.title('Form for Fun')

tabControl = ttk.Notebook(newForm)

# CREATE TAB CONTAINER
tab1 = ttk.Frame(tabControl)
tabControl.add(tab1, text='Container')

# CREATE TAB BOOKING
tab2 = ttk.Frame(tabControl)
tabControl.add(tab2, text='Booking')
tabControl.pack(expand=1, fill='both')

frame1 = Frame(tab1, borderwidth=2, relief="flat", padx=5,pady=5)
frame2 = Frame(tab1, relief="flat", padx=5, pady=20)
frame3 = Frame(tab1, borderwidth=2, relief="flat", padx=5, pady=5)

lbl1 = Label(frame1, text='QUẢN LÝ CONTAINER', bg='light blue', font='none 18 bold', relief='groove').grid(sticky="NESW")

lblSoContainer1 = Label(frame2, text='Số Container').grid(row=0, column=0, sticky=W, ipadx=20, ipady=10)
lblLoaiContainer1 = Label(frame2, text='Loại Container').grid(row=1, column=0, sticky=W, ipadx=20, ipady=10)
lblStatus1 = Label(frame2, text='Tình Trạng').grid(row=2, column=0, sticky=W, ipadx=20, ipady=10)
lblVaoCang1 = Label(frame2, text='Ngày Vào Cảng').grid(row=0, column=2, sticky=W, ipadx=20, ipady=10)
lblIE1 = Label(frame2, text="Xuất/Nhập Khẩu").grid(row=1, column=2, sticky=W, ipadx=20, ipady=10)
lblCongTy1 = Label(frame2, text="Công Ty").grid(row=2, column=2, sticky=W, ipadx=20, ipady=10)

soContainer1 = Entry(frame2, width=20)
soContainer1.grid(row=0, column=1)

lblExcel = Label(frame2, text ="Xuất Excel không thành công!",fg='red')
lblExcel.grid(row=5, column=1,sticky=W)
lblExcel.grid_remove()

# Fetch data for LOAI_CONTAINER dropdown box
cursor.execute(""" SELECT NAME FROM COM_CODE WHERE CODE LIKE '001%'""")
tableLoai = cursor.fetchall()
listLoai = []
for row in tableLoai:
    listLoai.append(row[0])
loaiContainer1 = ttk.Combobox(frame2, width=20, values=listLoai)
loaiContainer1.grid(row=1, column=1)

# Fetch data for TINH_TRANG dropdown box
cursor.execute(""" SELECT NAME FROM COM_CODE WHERE CODE LIKE '002%'""")
tableStatus = cursor.fetchall()
listStatus = []
for row in tableStatus:
    listStatus.append(row[0])
status1 = ttk.Combobox(frame2, width=20, values=listStatus)
status1.grid(row=2, column=1)

namVaoCang1 = ttk.Combobox(frame2, width=5, values=list(range(2015, 2026)))
namVaoCang1.grid(row=0, column=3)
thangVaoCang1 = ttk.Combobox(frame2, width=3, values=list(range(1, 13)))
thangVaoCang1.grid(row=0, column=4)
ngayVaoCang1 = ttk.Combobox(frame2, width=3, values=list(range(1, 32)))
ngayVaoCang1.grid(row=0, column=5)

# Fetch data for IE dropdown box
cursor.execute(""" SELECT NAME FROM COM_CODE WHERE CODE LIKE '003%'""")
tableIE = cursor.fetchall()
listIE = []
for row in tableIE:
    listIE.append(row[0])
xuatNhapKhau1 = ttk.Combobox(frame2, width=20, values=listIE)
xuatNhapKhau1.grid(row=1, column=3, columnspan=3)

# Fetch data for CONG_TY
cursor.execute('SELECT CONG_TY FROM CONTAINER')
tableCompany = cursor.fetchall()
listCompany = []
for row in tableCompany:
    listCompany.append(row[0])
congTy1 = ttk.Combobox(frame2, width=20, values=list(set(listCompany)))
congTy1.grid(row=2, column=3, columnspan=3)

# BUTTONS
btnThem1 = Button(frame3, text='Thêm', width=12, command=insertIntoTab1).grid(row=0, column=5, padx=(20, 0))
btnLuu1 = Button(frame3, text='Lưu', width=12, command=updateTab1).grid(row=1, column=5, padx=(20, 0))
btnXoa1 = Button(frame3, text='Xóa', width=12, command=deleteFlagTab1).grid(row=2, column=5, padx=(20, 0))
btnXuatExcel1 = Button(frame3, text='Xuất Excel', width=12, command=exportExcelTab1).grid(row=3, column=5, padx=(20, 0))

# CREATE CONTAINER TABLE
listColsTab1 = ('Số container', 'Loại container', 'Tình trạng', 'Ngày vào cảng', 'Xuất/Nhập', 'Công ty','Ngày tạo')
listBox1 = ttk.Treeview(frame3, columns=listColsTab1, show='headings')
for col in listColsTab1:
    listBox1.heading(col, text=col)
    listBox1.column(col,width=100)
listBox1.grid(row=0, column=0, columnspan=5, rowspan=5)
listBox1.bind('<<TreeviewSelect>>',fillEntryTab1)

frame1.pack()
frame2.pack()
frame3.pack()


# BOOKING WINDOW
frame4 = Frame(tab2, borderwidth=3, relief="flat", padx=5,pady=5)
frame5 = Frame(tab2, relief="flat", padx=5, pady=20)
frame6 = Frame(tab2, borderwidth=2, relief="flat", padx=5, pady=5)

lbl2 = Label(frame4, text='QUẢN LÝ ĐẶT CHỖ CONTAINER', bg='light blue', font='none 18 bold', relief='groove').grid()


lblSoBooking2 = Label(frame5, text='Số Booking').grid(row=0, column=0, sticky=W, ipadx=20, ipady=10)
lblSoContainer2 = Label(frame5, text='Số container').grid(row=1, column=0, sticky=W, ipadx=20, ipady=10)
lblVaoCang2 = Label(frame5, text='Ngày Vào Cảng').grid(row=0, column=3, sticky=W, ipadx=20, ipady=10)
lblCongTy2 = Label(frame5, text="Công Ty").grid(row=1, column=3, sticky=W, ipadx=20, ipady=10)

soBooking2 = Entry(frame5, width=35)
soBooking2.grid(row=0, column=1, columnspan=2)
soContainer2 = Entry(frame5, width=20)
soContainer2.grid(row=1, column=1)

namVaoCang2 = ttk.Combobox(frame5, width=5, values=list(range(2015, 2026)))
namVaoCang2.grid(row=0, column=4)
thangVaoCang2 = ttk.Combobox(frame5, width=3, values=list(range(1, 13)))
thangVaoCang2.grid(row=0, column=5)
ngayVaoCang2 = ttk.Combobox(frame5, width=3, values=list(range(1, 32)))
ngayVaoCang2.grid(row=0, column=6)

labelWarning = Label(frame5, text ="**Số Container không hợp lệ",fg='red')
labelWarning.grid(row=5, column=1,sticky=W)
labelWarning.grid_remove()

# Fetch data for CONG_TY BOOKING
cursor.execute('SELECT CONG_TY FROM BOOKING')
tableCompany2 = cursor.fetchall()
listCompany2 = []
for row in tableCompany2:
    listCompany2.append(row[0])
congTy2 = ttk.Combobox(frame5, width=20, values=list(set(listCompany2)))
congTy2.grid(row=1, column=4, columnspan=3)

# BOOKING BUTTONS
btnCheck2=Button(frame5, text='Check', width=7, command=checkBtn).grid(row=1, column=2,padx=5,pady=5)
btnThem2 = Button(frame6, text='Thêm', width=12, command=insertIntoTab2).grid(row=0, column=5, padx=(20, 0))
btnLuu2 = Button(frame6, text='Lưu', width=12, command=updateTab2).grid(row=1, column=5, padx=(20, 0))
btnXoa2 = Button(frame6, text='Xóa', width=12, command=deleteRowTab2).grid(row=2, column=5, padx=(20, 0))
btnXuatExcel2 = Button(frame6, text='Xuất Excel', width=12, command=exportExcelTab2).grid(row=3, column=5, padx=(20, 0))

# CREATE BOOKING TABLE
listColsTab2=('Số Booking', 'Số container', 'Ngày vào cảng', 'Công ty','Ngày tạo')
listBox2 = ttk.Treeview(frame6, columns=listColsTab2, show='headings')
for col in listColsTab2:
    listBox2.heading(col, text=col)
    listBox2.column(col,width=120)
listBox2.grid(row=0, column=0, columnspan=5, rowspan=5)
listBox2.bind('<<TreeviewSelect>>',fillEntryTab2)

frame4.pack()
frame5.pack()
frame6.pack()

showdataTab1()
showdataTab2()

newForm.mainloop()
