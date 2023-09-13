from win32com import client
from docxtpl import DocxTemplate
from docx.shared import Inches
from PyQt6.QtWidgets import QApplication, QMainWindow, QDialog, QFileDialog
from PyQt6 import uic, QtGui
from os import listdir
from os.path import isfile, join
import shutil
import os
import openpyxl
import datetime
import barcode_generator
import sys
import codecs
import ctypes
import math


def loadConfig():
    global fpath
    f = codecs.open("config.txt", "r", "utf-8")
    fpath = f.read()
    f.close()
    if fpath.isspace() or fpath == "":
        fpath=os.path.join(os.getcwd(), "receipts") 


def populate_excel(wb_path, c_data):
    excel = openpyxl.load_workbook(wb_path) 
    sheet = excel.active 
    sheet["A1"] = c_data["name"]
    sheet["A2"] = c_data["address"]
    sheet["A3"] = c_data["postalcode"]
    sheet["G6"] = c_data["refrence"]
    sheet["G7"] = c_data["date"].strftime("%d.%m.%Y")
    sheet["G8"] = c_data["date"].strftime("%H:%M")
    sheet["A17"] = c_data["description_receipt"]
    sheet["G21"] = c_data["amount_notax"]

    if c_data["OIB"] != "":
        sheet["A4"] = ("OIB:" + c_data["OIB"])

    excel.save(wb_path)
    excel.close()


loadConfig()
costumer_data = {
    "amount_notax" : 800, #Amount without tax
    "currency" : "EUR", #HRK/EUR
    "name" : "", # Name Surname
    "address" : "", #Full address, #
    "postalcode" : "", #Postal code, City
    "receiptID" : 0, #ID racuna
    "refrence" : "12345", #broj računa
    "description" : "", #kratki opis na uplati
    "date" : datetime.datetime.now(),
    "description_receipt" : "", #dugi opis na računu
    "amount_tax" : 1000, #Amount after tax
    "OIB" : "" #Personal identification number
}


class ConfirmDialog(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi("GUI/diag.ui", self)
        self.okbutton.clicked.connect(self.close)

class settingsWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi("GUI/settings.ui", self)
        self.accepted.connect(self.setFolderPath)
        self.browse_btn.clicked.connect(self.ask_filename)

    def setFolderPath(self):
        global fpath
        fpath = self.folder_edit.text()
        f = codecs.open("config.txt", "w", "utf-8")
        f.write(fpath)
        f.close()


    def ask_filename(self):
        fname = QFileDialog.getExistingDirectory(self)
        self.folder_edit.setText(fname)
    

class Window(QMainWindow):
    
    def __init__(self):
        super(Window, self).__init__()
        uic.loadUi("GUI/app.ui", self)
        self.setWindowIcon(QtGui.QIcon('icons/logo.png'))
        myappid = 'Companyname.receiptgenerator' # arbitrary string
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        self.IDwarning.setHidden(True)
        self.generate_button.clicked.connect(self.getData)
        self.calc_btn.clicked.connect(lambda: self.checkReceiptID(flag=True))
        self.amount_notax.editingFinished.connect(self.UpdateTax)
        self.amount_tax.editingFinished.connect(self.UpdateNoTax)
        self.receiptID.textChanged.connect(self.checkReceiptID)
        self.Settings.triggered.connect(self.settingsBox)
        

    def getData(self):
        costumer_data["name"] = self.name.text()
        costumer_data["amount_notax"] = float(self.amount_notax.text())
        costumer_data["currency"] = str(self.currency.currentText())
        costumer_data["address"] = self.address.text()
        costumer_data["postalcode"] = self.postalcode.text()
        costumer_data["receiptID"] = self.receiptID.text()
        costumer_data["refrence"] = costumer_data["receiptID"] + "-2-" + costumer_data["date"].strftime("%y")
        costumer_data["description"] = self.description.text()
        costumer_data["description_receipt"] = self.description_receipt.toPlainText()
        costumer_data["amount_tax"] = float(self.amount_tax.text())
        costumer_data["OIB"] = self.oib.text()
        self.generateReceipt()


    def dialogbox(self):
        #self.hide()
        self.myDialog = ConfirmDialog()
        self.myDialog.show()


    def settingsBox(self):
        self.settingsWindow = settingsWindow()
        self.settingsWindow.show()


    def UpdateTax(self):
        try:
            amount = float(self.amount_notax.text()) * 1.25 ##Change tax if needed
            amount = math.ceil(amount * 100)/100
            self.amount_tax.setText(str(amount))
        except ValueError:
            self.amount_tax.setText("")
            return
    def UpdateNoTax(self):
        try:
            amount = float(self.amount_tax.text()) * 0.8 ##Change tax if needed
            amount = math.ceil(amount * 100)/100
            self.amount_notax.setText(str(amount))
        except ValueError:
            self.amount_notax.setText("")
            return


    def checkReceiptID(self, flag=False):
        global fpath
        onlyfiles = [f for f in listdir(fpath) if isfile(join(fpath, f))]
        identifier = "-2-"+ costumer_data["date"].strftime("%y")
        templist=[]
        ids=[]
        for y in range(len(onlyfiles)):
            if identifier in onlyfiles[y]:
                templist.append(onlyfiles[y])
            else:
                pass

        for x in range(len(templist)):
            ids.append(templist[x][templist[x].find("."):])
            ids[x] = (ids[x][:ids[x].find("-")])
            ids[x] = int((ids[x][2:]))


        if flag == True:
            if ids:
                self.receiptID.setText(str(max(ids)+1))
            else:
                self.receiptID.setText("1")
        else:
            try:
                if int(self.receiptID.text()) in ids:
                    self.IDwarning.setHidden(False)
                else:
                    self.IDwarning.setHidden(True)
            except ValueError:
                self.IDwarning.setHidden(True)
        
        


    def generateReceipt(self):
        dirpath= os.getcwd()
        wordpath=os.path.join(dirpath,"templates\\template.docx")
        excelpath=os.path.join(dirpath,"templates\\template.xlsx")
        newwordpath=os.path.join(dirpath,"templates\\result.docx")
        newexcelpath=os.path.join(dirpath,"templates\\result.xlsx")

        #Make new files
        shutil.copyfile(wordpath, newwordpath)
        shutil.copyfile(excelpath, newexcelpath)

        #Get barcode img and fill excel file with costumer_data
        barcode_generator.generate(costumer_data)
        populate_excel(newexcelpath, costumer_data)

        ### Excel to Word copy-paste
        excel = client.Dispatch("Excel.Application")
        word = client.Dispatch("Word.Application")
        print(newwordpath)
        doc = word.Documents.Open(newwordpath)
        book = excel.Workbooks.Open(newexcelpath)
        sheet = book.Worksheets(1)
        sheet.Range("A1:G35").Copy()      # Selected the table I need to copy
        wdRange = doc.Content
        wdRange.Collapse(1) #start of the document, use 0 for end of the document
        wdRange.PasteExcelTable(False, False, False)
        sheet.Range("A1:A1").Copy() #to disable warning for full clipboard on close
        book.Close(True)
        doc.Close(True)

        ### This inserts barcode image to file, I haven't found simpler method for this.
        tpl = DocxTemplate(newwordpath)
        sd = tpl.new_subdoc()
        sd.add_picture("barcode\\" 
                        + costumer_data["refrence"] 
                        + costumer_data["name"].replace(" ","") 
                        + ".jpg", 
                        width=Inches(2.948), height=Inches(0.771))
        context = {'mysubdoc' : sd}
        tpl.render(context)
        saveloc = fpath + "\\receipt "+ costumer_data["date"].strftime("%Y") + ". " + costumer_data["refrence"]  + ", " + costumer_data["name"] + ".docx"
        print(saveloc)
        tpl.save(saveloc)

        #cleanup temp files
        os.remove(newwordpath)
        os.remove(newexcelpath)
        #self.dialogbox()
        os.startfile(saveloc)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    GUI = Window()
    GUI.show()
    sys.exit(app.exec())



