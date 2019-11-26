# -*- coding: utf-8 -*-

from tkinter import *
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
from tkinter import filedialog as fd


if __name__ == "__main__":
    techpath = 'tech.txt'
    grouppath = "group.txt"
    peoplepath = 'people.txt'

document = Document()
style = document.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)


def textfieldtodocx(event):
    line = textField.get(2.0, END)
    mylist = line.split('\n')
    mylist.pop()
    for i in range(len(mylist)):
        mylist[i] = str(i+1)+'|'+str(mylist[i])
    records = []
#    tmplist = []
    for x in range(len(mylist)):
        tmplist = mylist[x].split('|')
        tmplist[2] = tmplist[2] + '\n(' + tmplist[3] + ')'
        del tmplist[3]
        records.append(tmplist)
    actpar = document.add_paragraph('Акт о переносе техники №' + actNumber.get())
    actpar.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    table = document.add_table(rows=1, cols=5, style='Table Grid')
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    table.autofit = False
    table.allow_autofit = False
    table.columns[0].width = Inches(0.4)      # № п/п
    table.columns[1].width = Inches(2.1)    # Перемещение
    table.columns[2].width = Inches(1.3)    # Инв. номер
    table.columns[3].width = Inches(1.3)    # Забрали из отдела\n ФИО
    table.columns[4].width = Inches(1.3)    # Поставили в отдел\n ФИО
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '№'
    hdr_cells[1].text = 'Перемещение'
    hdr_cells[2].text = 'Инв. номер'
    hdr_cells[3].text = 'Забрали из отдела\n ФИО'
    hdr_cells[4].text = 'Поставили в отдел\n ФИО'

    for num, ids, invent, out, inp in records:
        row_cells = table.add_row().cells
        row_cells[0].text = num
        row_cells[1].text = ids
        row_cells[2].text = invent
        row_cells[3].text = out
        row_cells[4].text = inp

    datepar = document.add_paragraph('\n\n' + dateEntry.get()+'			Подпись_____________		'
                                     + people.get())
    datepar.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    file_name = fd.asksaveasfilename(initialfile=('#'+actNumber.get()+'_'+dateEntry.get()),
                                     defaultextension=".docx", filetypes=[('docx files', '*.docx')])
    document.save(file_name)
    root.destroy()


def nextvar(event):
    textField.config(state=NORMAL)
    textField.insert(END, '\n' + tech.get() + '|' + inventNumber.get() + '|' + codeNumber.get() + '|' +
                     groupOut.get() + '|' + groupIn.get())
    textField.config(state=DISABLED, bg="gainsboro")


def edittextfield(event):
    textField.config(state=NORMAL, bg="white")


techtxt = open(techpath, 'r', encoding="utf-8")
techlist = techtxt.read().splitlines()
grouptxt = open(grouppath, 'r', encoding="utf-8")
grouplist = grouptxt.read().splitlines()
peopletxt = open(peoplepath, 'r', encoding="utf-8")
peoplelist = peopletxt.read().splitlines()

root = Tk()
root.title('Mover 9000')
root.resizable(False, False)

numberFrame = Frame(root)
numberFrame.pack()
actNumber = StringVar(root)
actLabel = Label(numberFrame, text='Акт о переносе техники №', font='Arial 16')
actLabel.pack(side=LEFT)
actNumberEntry = Entry(numberFrame, font='Arial 16', justify='center', textvariable=actNumber, width=4)
actNumberEntry.pack(side=LEFT)
# ---------------------------BUTTONS-----------------------------------------------------------------------------------
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
inventNumberEntry = Entry(buttonFrame, font='Arial 16', justify='center', textvariable=inventNumber, width=15)
codeNumberEntry = Entry(buttonFrame, font='Arial 16', justify='center', textvariable=codeNumber, width=10)
groupOutMenu = OptionMenu(buttonFrame, groupOut, *grouplist)
groupOutMenu.config(width=21, font=12)
groupInMenu = OptionMenu(buttonFrame, groupIn, *grouplist)
groupInMenu.config(width=21, font=12)
buttonPlus = Button(buttonFrame, text=' + ', font='Arial 16', justify='center')
buttonPlus.bind('<Button-1>', nextvar)
techMenu.pack(side=LEFT)
inventNumberEntry.pack(side=LEFT)
codeNumberEntry.pack(side=LEFT)
groupOutMenu.pack(side=LEFT)
groupInMenu.pack(side=LEFT)
buttonPlus.pack(side=LEFT)
# --------------------------BUTTONS--END-------------------------------------------------------------------------------

textFrame = Frame(root)
textFrame.pack()
textField = Text(textFrame, width=150, state=DISABLED, bg="gainsboro")
scrollText = Scrollbar(textFrame, command=textField.yview)
scrollText.pack(side=RIGHT, fill=Y)
textField.config(yscrollcommand=scrollText.set)
textField.pack()
doneFrame = Frame(root)
doneFrame.pack(fill=X)
editButton = Button(doneFrame, text='Редактировать список', font=16)
doneButton = Button(doneFrame, text='ГОТОВО', font=18)
editButton.pack(side=TOP, fill=X)
editButton.bind('<Button-1>', edittextfield)
nLabel = Label(doneFrame, text='\n', font='Arial 2')
nLabel.pack(fill=X)
dateFrame = Frame(doneFrame)
dateFrame.pack(fill=X)

noow = datetime.now()
docDate = StringVar()
dateEntry = Entry(dateFrame, font='Arial 16', justify='center', textvariable=docDate)
dateEntry.insert(0, str(noow.day)+'.'+str(noow.month)+'.'+str(noow.year))
dateEntry.pack(side=LEFT)

people = StringVar(root)
people.set(peoplelist[0])
peopleMenu = OptionMenu(dateFrame, people, *peoplelist)
peopleMenu.config(width=21, font=12)
peopleMenu.pack(side=RIGHT)

n2Label = Label(doneFrame, text='\n', font='Arial 2')
n2Label.pack(fill=X)
doneButton.pack(fill=X)
doneButton.bind('<Button-1>', textfieldtodocx)

root.mainloop()
