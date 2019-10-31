## -*- coding: utf-8 -*-

from tkinter import *
import tkinter.ttk as ttk
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT


document = Document()
style = document.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)


def textFieldToDocx(event):
    line = textField.get(2.0, END)
    mylist = line.split('\n')
    mylist.pop()
    for i in range(len(mylist)):
        mylist[i] = str(i+1)+'|'+str(mylist[i])
    records = []
    tmplist = []
    for x in range(len(mylist)):
        tmplist = mylist[x].split('|')
        tmplist[2] = tmplist[2] + '\n(' + tmplist[3] + ')'
        del tmplist[3]
        records.append(tmplist)

    table = document.add_table(rows=1, cols=5, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    table.autofit = False
    table.allow_autofit = False
    table.columns[0].width = Cm(0.7)    # № п/п
    table.columns[1].width = Cm(6.1)    # Перемещение
    table.columns[2].width = Cm(4.0)    # Инв. номер
    table.columns[3].width = Cm(3.6)    # Забрали из отдела\n ФИО
    table.columns[4].width = Cm(3.6)    # Поставили в отдел\n ФИО
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '№'
    hdr_cells[1].text = 'Перемещение'
    hdr_cells[2].text = 'Инв. номер'
    hdr_cells[3].text = 'Забрали из отдела\n ФИО'
    hdr_cells[4].text = 'Поставили в отдел\n ФИО'

    for num, id, invent, out, inp in records:
        row_cells = table.add_row().cells
        row_cells[0].text = num
        row_cells[1].text = id
        row_cells[2].text = invent
        row_cells[3].text = out
        row_cells[4].text = inp

    document.save('test.docx')
    print('ГОТОВО')


def nextVar(event):
    textField.config(state=NORMAL)
    textField.insert(END,'\n' + tech.get() + '|' + inventNumber.get() + '|' + codeNumber.get() + '|' +
                     groupOut.get() + '|' + groupIn.get())
    textField.config(state=DISABLED, bg="gainsboro")


def editTextField(event):
    textField.config(state=NORMAL, bg="white")


techtxt = open('tech.txt','r',encoding="utf-8")
techlist = techtxt.read().splitlines()
grouptxt = open('group.txt','r',encoding="utf-8")
grouplist = grouptxt.read().splitlines()

root = Tk()
root.title('Mover 9000')
root.resizable(False, False)

#----------------------------BUTTONS-----------------------------------------------------------------------------------
buttonFrame = Frame(root)
buttonFrame.pack()

tech = StringVar(root)
tech.set(techlist[0])
groupOut = StringVar(root)
groupIn = StringVar(root)
groupOut.set(grouplist[0])
groupIn.set(grouplist[1])
inventNumber = StringVar(root)
codeNumber = StringVar(root)
techMenu = OptionMenu(buttonFrame, tech, *techlist)
techMenu.config(width=35, font=12)
inventNumberEntry = Entry(buttonFrame,font='Arial 16', justify='center', textvariable=inventNumber)
codeNumberEntry = Entry(buttonFrame, font='Arial 16', justify='center', textvariable=codeNumber, width=10)
groupOutMenu = OptionMenu(buttonFrame, groupOut, *grouplist)
groupOutMenu.config(width=21, font=12)
groupInMenu = OptionMenu(buttonFrame, groupIn, *grouplist)
groupInMenu.config(width=21, font=12)
buttonPlus = Button(buttonFrame, text=' + ', font='Arial 16', justify='center')
buttonPlus.bind('<Button-1>',nextVar)

techMenu.pack(side=LEFT)
inventNumberEntry.pack(side=LEFT)
codeNumberEntry.pack(side=LEFT)
groupOutMenu.pack(side=LEFT)
groupInMenu.pack(side=LEFT)
buttonPlus.pack(side=LEFT)
#---------------------------BUTTONS--END-------------------------------------------------------------------------------

textFrame = Frame(root)
textFrame.pack()
textField = Text(textFrame,width=150, state=DISABLED,bg="gainsboro")
scrollText = Scrollbar(textFrame, command=textField.yview)
scrollText.pack(side=RIGHT, fill=Y)
textField.config(yscrollcommand=scrollText.set)
textField.pack()

doneFrame = Frame(root)
doneFrame.pack(fill=X)
editButton = Button(doneFrame, text = 'РЕДАКТИРОВАТЬ')
doneButton = Button(doneFrame, text ='ГОТОВО', font=18)
editButton.pack(side=TOP, fill=X)
doneButton.pack(fill=X)
doneButton.bind('<Button-1>',textFieldToDocx)
editButton.bind('<Button-1>', editTextField)

root.mainloop()