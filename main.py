from tkinter import *
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from tkinter import font as tkfont
from datetime import date
import time

#FUNGSI Tk,Openpyxl
root = Tk()
root.title("absensi")
root.resizable(width=False, height=False)
workbook = Workbook()
sheet = workbook.active

#FONT
styling = tkfont.Font(family='Helvatica', weight='bold', size=20)
styling2 = tkfont.Font(family='helvatica', size=9)
font = Font(bold=True)
border = Border(left=Side(border_style='thin', color='00000000'),
                right=Side(border_style='thin', color='00000000'),
                top=Side(border_style='thin', color='00000000'),
                bottom=Side(border_style='thin', color='00000000'))
alignment = Alignment(horizontal='center', vertical='center')

#WAKTU
today=date.today()
time_string = time.strftime('%H:%M:%S')
now = time.strftime('%H%M%S')
def display_time():
    global time_string, now
    time_string = time.strftime('%H:%M:%S')
    now = time.strftime('%H%M%S')
    timenow['text'] = time_string
    root.after(1000, display_time)

#MAIN SCREEN
height = 500
width = 600
canvas = Canvas(root, height=height, width=width, bg="lightblue")
canvas.pack()

#AWALAN EXCEL
sheet['A1'] = 'Date :'
A1 = sheet['A1']
A1.font = font

sheet['B1'] = str(today)
A1 = sheet['A2']
A1.font = font

sheet['A2'] = 'Data Absensi Pegawai BC'
A1 = sheet['A2']
A1.font = font

sheet['A3'] = "No"
A3 = sheet['A3']
A3.font = font
A3.border = border
A3.alignment = alignment

sheet['B3'] = "Nama"
B3 = sheet['B3']
B3.font = font
B3.border = border
B3.alignment = alignment

sheet['C3'] = "NIP"
C3 = sheet['C3']
C3.font = font
C3.border = border
C3.alignment = alignment

sheet['D3'] = "Time"
D3 = sheet['D3']
D3.font = font
D3.border = border
D3.alignment = alignment

#VARIABLE JUMLAH LOGIN
num = 0

#FUNGSI TOMBOL INSERT
def insertdata():
    global num
    num = num+1

    shetnum = num + 3
    sheet['A'+str(shetnum)] = num
    datano = sheet['A'+str(shetnum)]
    datano.border = border
    datano.alignment = alignment

    sheet['B' + str(shetnum)] = NamaEntry.get()
    datanama = sheet['B' + str(shetnum)]
    datanama.border = border
    datanama.alignment = alignment

    sheet['C' + str(shetnum)] = NIPEntry.get()
    dataNIP= sheet['C' + str(shetnum)]
    dataNIP.border = border
    dataNIP.alignment = alignment

    sheet['D' + str(shetnum)] = str(time_string)
    datatime = sheet['D' + str(shetnum)]
    datatime.border = border
    datatime.alignment = alignment

    informasi['text'] = ('Data absensi ' + str(NamaEntry.get()) + ' telah dimasukkan!'+' Tekan save, jika sudah selesai!')
    NamaEntry.delete(0, END)
    NIPEntry.delete(0, END)
    save['state'] = 'normal'
    exitt['state'] = 'disabled'

#FUNGSI TOMBOL SAVE
def savedata():
    global informasi
    workbook.save(filename=str(today)+str(now)+".xlsx")
    informasi['text'] = ('Data absen telah disave pada file '+str(today)+str(now)+".xlsx,"+" Klik 'NEW' untuk menginput ulang !")
    NamaEntry['state'] = 'disabled'
    NIPEntry['state'] = 'disabled'
    save['state'] = 'disabled'
    insert['state'] = 'disabled'
    ClearNama['state'] = 'disabled'
    ClearNIP['state'] = 'disabled'
    new['state'] = 'normal'
    exitt['state'] = 'normal'

#FUNGSI TOMBOL NEW
def create():
    global informasi, num
    informasi['text'] = 'Berhasil membuat excel baru, Silahkan input kembali'
    NamaEntry.delete(0, END)
    NIPEntry.delete(0, END)
    NamaEntry['state'] = 'normal'
    NIPEntry['state'] = 'normal'
    save['state'] = 'normal'
    insert['state'] = 'normal'
    ClearNama['state'] = 'normal'
    ClearNIP['state'] = 'normal'
    num = 0

#FUNGSI TOMBOL CLEAR
def clearN():
    NamaEntry.delete(0,END)

def clearID():
    NIPEntry.delete(0, END)

#JUDUL
framejudul = Frame(root, bg='white')
framejudul.place(rely=0.1, relx=0.05, relheight=0.1, relwidth=0.4)
ljudul = Label(framejudul, bg='white', text='Absensi Pegawai', font=styling)
ljudul.place(relheight=1, relwidth=1)

#DISPLAY TIME
frametime = Frame(root, bg='white')
frametime.place(rely=0.25, relx=0.05, relheight=0.1, relwidth=0.3)
timenow = Label(frametime, font='ariel 30', fg='black')
timenow.place(relheight=1, relwidth=1)
display_time()

#DISPLAY NAMA
frameNama = Frame(root, bg='white')
frameNama.place(rely=0.4, relx=0.5, relheight=0.06, relwidth=0.9, anchor='n')
Namainfo = Label(frameNama, bg='white', text='Nama', font=styling2)
Namainfo.place(relheight=1, relwidth=0.2)
NamaEntry = Entry(frameNama)
NamaEntry.place(relx=0.2, relheight=1, relwidth=0.8)
NamaEntry.get()
ClearNama = Button(frameNama, text='Clear', command=clearN)
ClearNama.place(relx=0.8, relheight=1, relwidth=0.2)

#DISPLAY ID
frameNIP = Frame(root, bg='white')
frameNIP.place(rely=0.5, relx=0.5, relheight=0.06, relwidth=0.9, anchor='n')
NIPinfo = Label(frameNIP, bg='white', text='NIP', font=styling2)
NIPinfo.place(relheight=1, relwidth=0.2)
NIPEntry = Entry(frameNIP)
NIPEntry.place(relx=0.2, relheight=1, relwidth=0.8)
NIPEntry.get()
ClearNIP= Button(frameNIP, text='Clear', command=clearID)
ClearNIP.place(relx=0.8, relheight=1, relwidth=0.2)

#DISPLAY KETERANGAN
informasi = Label(root, bg='white', text='Informasi', font=styling2)
informasi.place(rely=0.6, relx=0.5, relheight=0.06, relwidth=0.9, anchor='n')

#DISPLAY BUTTON
frameSUB = Frame(root, bg='white')
frameSUB.place(rely=0.7, relx=0.5, relheight=0.06, relwidth=0.3, anchor='n')
insert = Button(frameSUB, text='INSERT', command=insertdata)
insert.place(relheight=1, relwidth=1/3)
save = Button(frameSUB, text='SAVE',state = 'disabled', command=savedata)
save.place(relx=1/3, relheight=1, relwidth=1/3)
new = Button(frameSUB, text='NEW',state = 'disabled' , command=create)
new.place(relx=2/3, relheight=1, relwidth=1/3)
exitt = Button(root, text='EXIT', command=root.quit)
exitt.place(rely=0.8, relx=0.5, relheight=0.06, relwidth=0.2, anchor='n')

root.mainloop()