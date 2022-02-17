import os
import win32com.client
import shutil
from tkinter import *

folders = next(os.walk("."))[1]
path = os.getcwd()
root = Tk()
invoiceDict = {}

myLabel = Label(root, text="Rechnungen schreiben")
myLabel.grid(row=0, column=0, columnspan=2)

invoiceNumberLabel = Label(root, text="Erste Rechnungsnummer:")
invoiceNumberLabel.grid(row=1, column=0, sticky="w")

invoiceNumberInput = Entry(root, width=30)
invoiceNumberInput.grid(row=1, column=1)

invoiceLabel = Label(root, text="Märkte auswählen:")
invoiceLabel.grid(row=2, column=0, sticky="w")

rowindex = 3
for folder in folders:
    invoiceDict["{}".format(folder)] = IntVar()
    checkbox=Checkbutton(root, text=folder, variable = invoiceDict[folder], onvalue = 1, offvalue = 0)
    checkbox.grid(row=rowindex, column=1, sticky="w")
    rowindex = rowindex + 1

def writeInvoices(invoiceNumber, folders):
    print("entered function writeInvoices")
    laufendeNummer = int(invoiceNumber.split('/')[1])
    prefix= int(invoiceNumber.split('/')[0])
    print("path: " + path)

    for folder in folders:
        files = os.listdir(path + "/" + folder)
        print("folder: " + folder)
        for file in files:
            if file.startswith("Summe"):
                filepath = path + "/" + folder + "/" + file
                print(filepath)

                standardText = "Rechnungsnummer: " + str(prefix) + "/" + str(laufendeNummer)            
                print(standardText)

                xl = win32com.client.Dispatch("Excel.Application")
                wb = xl.Workbooks.Open(filepath, ReadOnly=0)
                ws = wb.Worksheets("Tabelle1")

                ws.Cells(17,3).Value = standardText
                xl.Application.Run("Modul1.WerteausDateien_addieren")

                laufendeNummer = laufendeNummer + 1

                wb.Save()
                wb.Close()

                xl.Quit()


def moveLieferscheine(folders):
    
    for folder in folders:
        files = os.listdir(path + "/" + folder)
        for file in files:
            if(file.endswith("Lieferscheine")):
                lieferscheinFolder = file
                
        for file in files:
            if file.endswith(".xlsx") and not file.startswith("Summe"):
                filepath = path + "/" + folder + "/" + file
                destpath = path + "/" + folder + "/" + lieferscheinFolder
                shutil.move(filepath, destpath)

def fEnter():
    invoiceNumber = invoiceNumberInput.get()
    finalInvoices = []
    infobox = Tk()
    
    if invoiceNumber == "" or not "/" in invoiceNumber:
        myRechnungsnummer = Label(infobox, text="Bitte Rechnungsnummer in richtigem Format eingeben")
        myRechnungsnummer.pack()
        return
    
    myRechnungsnummer = Label(infobox, text=invoiceNumber)

    myRechnungsnummer.pack()
    
    for invoice in invoiceDict:
        if invoiceDict[invoice].get() == 1:
            finalInvoices.append(invoice)
            myInvoicesList = Label(infobox, text=invoice)
            myInvoicesList.pack()
            
            
    if finalInvoices:
        writeInvoices(invoiceNumber, finalInvoices)
        moveLieferscheine(finalInvoices)
    
    else:
        myInvoicesList = Label(infobox, text="Bitte Märkte auswählen")
        myInvoicesList.pack()
    
    cancelButton2 = Button(infobox, text="Ok", padx=50, pady=10, command=infobox.destroy)
    cancelButton2.pack()


    

cancelButton = Button(root, text="Abbrechen", padx=50, pady=10, command=root.destroy)
cancelButton.grid(row=rowindex, column=0)

enterButton = Button(root, text="Rechnungen schreiben", padx=50, pady=10, command=fEnter)
enterButton.grid(row=rowindex, column=1)


root.mainloop()




