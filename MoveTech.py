## -*- coding: utf-8 -*-

from tkinter import *
import tkinter.ttk as ttk


def ttget(event):
    line = textField.get(2.0, END)
    mylist = line.split('\n')
    for i in range(len(mylist)):
        mylist[i] = str(i+1)+'|'+str(mylist[i])
    print(mylist)


def nextVar(event):
    textField.config(state=NORMAL)
    textField.insert(END,'\n' + tech.get() + '|' + inventNumber.get() + '|' + codeNumber.get() + '|' +
                     groupOut.get() + '|' + groupIn.get())
    textField.config(state=DISABLED, bg="gainsboro")


def editTextField(event):
    textField.config(state=NORMAL, bg="white")


def textFieldToDocx(event):
    textTable = textField.get(2.0, END)
    print(textTable)


techtxt = open('tech.txt','r',encoding="utf-8")
techlist = techtxt.read().splitlines()
grouptxt = open('group.txt','r',encoding="utf-8")
grouplist = grouptxt.read().splitlines()

root = Tk()
root.title('Mover 9000')
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
inventNumberEntry = Entry(buttonFrame, justify='center',textvariable=inventNumber)
codeNumberEntry = Entry(buttonFrame, justify='center',textvariable=codeNumber)
groupOutMenu = OptionMenu(buttonFrame, groupOut, *grouplist)
groupInMenu = OptionMenu(buttonFrame, groupIn, *grouplist)
buttonPlus = Button(buttonFrame, text=' + ')
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
textField = Text(textFrame,width=100, state=DISABLED,bg="gainsboro")
scrollText = Scrollbar(textFrame, command=textField.yview)
scrollText.pack(side=RIGHT, fill=Y)
textField.config(yscrollcommand=scrollText.set)
textField.pack()

doneFrame = Frame(root)
doneFrame.pack(fill=X)
editButton = Button(doneFrame, text = 'РЕДАКТИРОВАТЬ')
doneButton = Button(doneFrame, text ='ГОТОВО')
editButton.pack(side=TOP, fill=X)
doneButton.pack(fill=X)
#doneButton.bind('<Button-1>',textFieldToDocx)
editButton.bind('<Button-1>', editTextField)

doneButton.bind('<Button-1>',ttget)

root.mainloop()