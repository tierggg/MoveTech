## -*- coding: utf-8 -*-

from tkinter import *
import tkinter.ttk as ttk

def printVar(event):
    textField.config(state=NORMAL)
    textField.insert(END, '\n' + inventNumber.get())
    textField.config(state=DISABLED)

texttxt = open('tech.txt','r',encoding="utf-8")
techlist = texttxt.read().splitlines()


root = Tk()
root.title('Mover 9000')

buttonFrame = Frame(root)
buttonFrame.pack()

tech = StringVar(root)
inventNumber = StringVar(root)
tech.set(techlist[0])
techMenu = OptionMenu(buttonFrame, tech, *techlist)
techMenu.pack(side=LEFT) #-----------------------------------------------------
inventNumberEntry = Entry(buttonFrame, justify='center',textvariable=inventNumber)
inventNumberEntry.pack(side=LEFT)

textFrame = Frame(root)
textFrame.pack()
textField = Text(textFrame, state = DISABLED)
textField.pack()

doneFrame = Frame(root)
doneFrame.pack()
doneButton = Button(doneFrame, text ='ГОТОВО')
doneButton.pack()
doneButton.bind('<Button-1>',printVar)


root.mainloop()