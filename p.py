import sys
import re
import time
from PyQt5.QtWidgets import QApplication, QWidget, QTableWidgetItem, QSizePolicy, QFileDialog, QDialog
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtCore import Qt
from ast import literal_eval
from pynput.mouse import Listener
import pyautogui
import pyperclip

ignoreItemsList = []
chooseColumnForTypeFlag = False
stratChooseColumnForMultiType = 0
actionDialog = ''

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('main.ui', self)
        
        self.actionDialog.hide()
        
        self.uploadPersonList.clicked.connect(self.loadFileWithHeader)
        
        self.setMouseClick.clicked.connect(self.setMouseClickCommand)
        self.typeOneColumnTextFromList.clicked.connect(self.setTypOneColumnTextFromList)
        self.typeMultiColumnTextFromList.clicked.connect(self.setTypMultiColumnTextFromList)
        self.typeCustomText.clicked.connect(self.showCustomTextPopup)
        self.setDelay.clicked.connect(self.showDelayPopup)
        self.moveUpCommand.clicked.connect(self.moveUpSelectedCommand)
        self.moveDownCommand.clicked.connect(self.moveDownSelectedCommand)
        self.removeCommand.clicked.connect(self.removeSelectedCommand)
        self.removeAllCommand.clicked.connect(self.restCommandList)
        
        self.commandSetCancel.clicked.connect(self.cancelOfCommand)
        self.commandSetEnd.clicked.connect(self.endOfCommand)
        

        self.stratProccess.clicked.connect(self.RunListofCommand)
        self.resetFileList()
        self.restCommandList()


    def loadFileWithHeader(self):
        fileName = QFileDialog.getOpenFileName(self, 'Choose File', r'C:', 'CSV files(*.csv)')
        if(fileName[0] != ''):
            self.resetFileList()
            lineCount = 0
            with open(fileName[0], 'r', encoding='UTF-8') as file:
                while (line := file.readline().rstrip()):
                    data = line.split(',')
                    
                    rowPosition = self.fileTable.rowCount()
                    self.fileTable.insertRow(rowPosition)
                    
                    if(lineCount == 0):
                        colCount = len(data) + 1
                        self.fileTable.setColumnCount(colCount)
                        
                        allCkeck = QTableWidgetItem('همه')
                        allCkeck.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                        allCkeck.setCheckState(Qt.CheckState.Checked)
                        self.fileTable.setItem(0, 0, allCkeck)
                        for col in range(1, colCount):
                            self.fileTable.setItem(0, col, QTableWidgetItem(data[col-1].strip()))
                            
                        self.fileTable.itemClicked.connect(self.tableItemClicked)    
                        header = self.fileTable.horizontalHeader()       
                        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
                        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
                        header.setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)                        
                    else:
                        for col in range(3):
                            if(col == 0):
                                item = QTableWidgetItem(str(lineCount))
                                item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                                item.setCheckState(Qt.CheckState.Checked)
                                self.fileTable.setItem(rowPosition, col, item)
                            else:
                                self.fileTable.setItem(rowPosition, col, QTableWidgetItem(data[col-1].strip())) 
                        
                        
                    lineCount += 1

    def tableItemClicked(self, item):
        global ignoreItemsList
        if(item.column() == 0):
            if(item.row() == 0):
                if item.checkState() == QtCore.Qt.Checked:
                    ignoreItemsList = []
                    for row in range(1, self.fileTable.rowCount()):
                        self.fileTable.item(row, 0).setCheckState(Qt.CheckState.Checked)
                    
                elif item.checkState() == QtCore.Qt.Unchecked:
                    ignoreItemsList = [i for i in range(1, self.fileTable.rowCount())]
                    for row in range(1, self.fileTable.rowCount()):
                        self.fileTable.item(row, 0).setCheckState(Qt.CheckState.Unchecked)
            else:
                if item.checkState() == QtCore.Qt.Checked:
                    if item.row() in ignoreItemsList:
                        ignoreItemsList.remove(item.row())
                elif item.checkState() == QtCore.Qt.Unchecked:
                    ignoreItemsList.append(item.row())

    def setMouseClickCommand(self):
        if actionDialog == '':
            self.showMinimized()
            def on_click(x, y, button, pressed):
                if pressed:
                    rowPosition = self.commandTable.rowCount()
                    self.commandTable.insertRow(rowPosition)
                    self.commandTable.setItem(rowPosition, 0, QTableWidgetItem('Mouse({}, {}, {})'.format(x, y, str(button).split('.')[-1]))) 
                else:
                    return False
                    
            with Listener(on_click=on_click) as listener:
                listener.join()

    def setTypOneColumnTextFromList(self):
        global chooseColumnForTypeFlag, actionDialog
        if actionDialog == '':
            self.setActionDialog("ستون مورد نظر را انتخاب نمایید.")
            actionDialog = 'chooseColumnForType'
            chooseColumnForTypeFlag = True
            self.fileTable.itemClicked.connect(self.chooseColumnForType)  

    def chooseColumnForType(self, item):
        global chooseColumnForTypeFlag
        if chooseColumnForTypeFlag:
            self.fileTable.itemClicked.disconnect(self.chooseColumnForType)  
            chooseColumnForTypeFlag = False
            rowPosition = self.commandTable.rowCount()
            self.commandTable.insertRow(rowPosition)
            self.commandTable.setItem(rowPosition, 0, QTableWidgetItem('ListType([\'{}\', {}])'.format(self.fileTable.item(0, item.column()).text(), item.column())))
        
    def setTypMultiColumnTextFromList(self):
        global chooseColumnForTypeFlag, actionDialog
        if actionDialog == '':
            self.setActionDialog("ستون های مورد نظر را انتخاب نمایید و سپس برروی دکمه اتمام دستور کلیک کنید.", 1)
            actionDialog = 'chooseMultiColumnForType'
            chooseColumnForTypeFlag = True
            self.fileTable.itemClicked.connect(self.chooseMultiColumnForType)

    def chooseMultiColumnForType(self, item):
        global stratChooseColumnForMultiType
        if chooseColumnForTypeFlag:
            rowPosition = self.commandTable.rowCount()
            lastRowCommand = self.commandTable.item(rowPosition - stratChooseColumnForMultiType, 0)
            if(lastRowCommand == None):
                stratChooseColumnForMultiType = 1
                self.commandTable.insertRow(rowPosition)
                commandContent = 'ListType([\'{}\', {}])'.format(self.fileTable.item(0, item.column()).text(), item.column())
                self.commandTable.setItem(rowPosition, 0, QTableWidgetItem(commandContent))
            else:
                commandContent = lastRowCommand.text().replace(')', '') + ', [\'{}\', {}])'.format(self.fileTable.item(0, item.column()).text(), item.column())
                self.commandTable.setItem(rowPosition - 1, 0, QTableWidgetItem(commandContent))

    def showCustomTextPopup(self):
        if actionDialog == '':
            customTextDialog = CustomTextPopup()

            if customTextDialog.exec_() == QtWidgets.QDialog.Accepted and customTextDialog.popupTextBox.toPlainText() != '':
                rowPosition = self.commandTable.rowCount()
                self.commandTable.insertRow(rowPosition)
                commandContent = 'CustomType({})'.format(customTextDialog.popupTextBox.toPlainText())
                self.commandTable.setItem(rowPosition, 0, QTableWidgetItem(commandContent))
            
    def showDelayPopup(self):
        if actionDialog == '':
            delayDialog = DelayPopup()

            if delayDialog.exec_() == QtWidgets.QDialog.Accepted and delayDialog.doubleSpinBox.text() != '0.00':
                rowPosition = self.commandTable.rowCount()
                self.commandTable.insertRow(rowPosition)
                commandContent = 'Delay({})'.format(delayDialog.doubleSpinBox.text())
                self.commandTable.setItem(rowPosition, 0, QTableWidgetItem(commandContent))

    def moveUpSelectedCommand(self): 
        if actionDialog == '':
            currentRow = self.commandTable.currentRow()
            if self.commandTable.item(currentRow,0).isSelected() and currentRow != 0:
                holdCurrent = self.commandTable.item(currentRow, 0).text()
                holdUp = self.commandTable.item(currentRow - 1, 0).text()
                self.commandTable.setItem(currentRow - 1, 0, QTableWidgetItem(holdCurrent))
                self.commandTable.setItem(currentRow, 0, QTableWidgetItem(holdUp))
                self.commandTable.setCurrentCell(currentRow - 1, 0)
        
    def moveDownSelectedCommand(self):
        if actionDialog == '':
            currentRow = self.commandTable.currentRow()
            if self.commandTable.item(currentRow,0).isSelected() and currentRow != (self.commandTable.rowCount() - 1):
                holdCurrent = self.commandTable.item(currentRow, 0).text()
                holdUp = self.commandTable.item(currentRow + 1, 0).text()
                self.commandTable.setItem(currentRow + 1, 0, QTableWidgetItem(holdCurrent))
                self.commandTable.setItem(currentRow, 0, QTableWidgetItem(holdUp))
                self.commandTable.setCurrentCell(currentRow + 1, 0)
            
    def removeSelectedCommand(self):
        if actionDialog == '':
            currentRow = self.commandTable.currentRow()
            if self.commandTable.item(currentRow, 0).isSelected():
                self.commandTable.removeRow(currentRow)
          
    def RunListofCommand(self):
        if(self.commandTable.rowCount() != 0):
            self.showMinimized()
            time.sleep(3)
            for row in range(1, self.fileTable.rowCount()):
                if(row not in ignoreItemsList):
                    for i in range(self.commandTable.rowCount()):
                        command = self.commandTable.item(i, 0).text()
                        cType = re.search('^(.+?)(?=\()', command).group(0)
                        cValue = re.search('(?s)(\(.*?\))(?=\s|$)', command).group(0).replace('(', '').replace(')', '')
                        
                        if cType == 'Mouse':
                            coordinate = cValue.split(', ')
                            pyautogui.click(int(coordinate[0]), int(coordinate[1]))
                        elif cType == 'ListType':
                            cText = ''
                            for index, item in enumerate(literal_eval(cValue)):
                                if isinstance(item, list):
                                    cText += self.fileTable.item(row, item[1]).text() + ' '
                                elif index == 1:
                                        cText += self.fileTable.item(row, item).text() + ' '
                                        
                            pyperclip.copy(cText.strip())
                            pyautogui.hotkey("ctrl", "v")
                        elif cType == 'CustomType':
                            pyperclip.copy(cValue.strip())
                            pyautogui.hotkey("ctrl", "v")
                        elif cType == 'Delay':
                            time.sleep(float(cValue))
                
    def setActionDialog(self, lable, end = 0):
        self.actionDialog.show()
        self.actionDialogLabel.setText(str(lable))
        if end == 0:
            self.commandSetEnd.hide()
        elif end == 1:
            self.commandSetEnd.show()

    def cancelOfCommand(self):
        global actionDialog
        self.actionDialog.hide()
        chooseColumnForTypeFlag = False
        if actionDialog == 'chooseMultiColumnForType':
            actionDialog = ''
            stratChooseColumnForMultiType = 0
            self.fileTable.itemClicked.disconnect(self.chooseMultiColumnForType)
        elif actionDialog == 'chooseColumnForType':
            actionDialog = ''
            self.fileTable.itemClicked.disconnect(self.chooseColumnForType)
        
    def endOfCommand(self):
        global actionDialog
        self.actionDialog.hide()
        actionDialog = ''
        stratChooseColumnForMultiType = 0
        chooseColumnForTypeFlag = False
        self.fileTable.itemClicked.disconnect(self.chooseMultiColumnForType)
    
    def resetFileList(self):
        ignoreItemsList = []
        self.fileTable.setRowCount(0)
        self.fileTable.setColumnCount(0)
       
    def restCommandList(self):
        if actionDialog == '':
            self.commandTable.setRowCount(0)
            self.commandTable.setColumnCount(1)
            chooseColumnForTypeFlag = False
            stratChooseColumnForMultiType = 0

class CustomTextPopup(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('customTextDialog.ui', self)
        self.confirmPopup.clicked.connect(self.accept)
        self.cancelPopup.clicked.connect(self.close)
        
    def accept(self):
        super().accept()

class DelayPopup(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('setDelayDialog.ui', self)
        self.confirmPopup.clicked.connect(self.accept)
        self.cancelPopup.clicked.connect(self.close)
        
    def accept(self):
        super().accept()
     
def main():
    app = QApplication(sys.argv)
    ex = MainWindow()
    ex.show()
    try:
        sys.exit(app.exec_())
    except:
        print("Closing Window...")
    

if __name__ == '__main__':
    main()