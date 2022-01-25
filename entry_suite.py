import datetime
import os
import sys  # for PyQt5.5 debugging
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import QtCore

form_class = uic.loadUiType("GUI/EntryForm.ui")[0]


class DataDNAIdentification:
    def __init__(self):
        self.analyst = ''
        self.date = datetime.datetime.now().strftime('%Y%m%d')

    def import_info(self):
        readfile_info = open('info.ini', mode='r')  # File I/O Error 가능성
        self.analyst = readfile_info.readline().rstrip('\n').split('=')[1]
        self.date = readfile_info.readline().rstrip('\n').split('=')[1]
        readfile_info.close()

    def export_info(self):
        writefile_info = open('info.ini', mode='w')     # File I/O Error 가능성
        writefile_info.write('analyst=%s\n' % self.analyst)
        writefile_info.write('date=%s\n' % self.date)
        writefile_info.close()

    def get_dir(self):
        return '%s_%s' % (self.date, self.analyst)


class EntryForm(QDialog, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        readfile_setting = open(os.getcwd() + '/Settings/Settings.ini', mode='r')  # File I/O Error 가능성
        location_save = readfile_setting.readline().rstrip('\n').split('=')[1]
        analyst = readfile_setting.readline().rstrip('\n').split('=')[1]
        readfile_setting.close()
        self.set_line_texts(location_save=location_save, analyst=analyst,
                            date=QtCore.QDate.currentDate())

    def set_line_texts(self, location_save="", analyst="", date=QtCore.QDate.currentDate(), optional=""):
        self.line_savelocation.setText(location_save)
        self.line_analyst.setText(analyst)
        self.line_date.setDate(date)
        self.line_optional.setText(optional)

    def set_line_ro(self, condition=True):
        self.line_analyst.setReadOnly(condition)
        self.line_date.setReadOnly(condition)
        self.line_optional.setReadOnly(condition)

    def click_btn_change_savelocation(self):
        foldername = QFileDialog.getExistingDirectory(self, 'Open folder')
        if foldername is not "":
            self.line_savelocation.setText(foldername)
        else:
            return

    def click_btn_new(self):
        QMessageBox.information(self, "Notice", "Input new profile and press OK button.")
        self.set_line_ro(False)

    def click_btn_load(self):
        location_load = QFileDialog.getExistingDirectory(self, 'Open folder')
        if location_load is not "":
            if os.path.exists('%s/%s' % (location_load, 'info.ini')):
                readfile_setting = open(location_load + '/info.ini', mode='r')  # File I/O Error 가능성
                analyst = readfile_setting.readline().rstrip('\n').split('=')[1]
                date = readfile_setting.readline().rstrip('\n').split('=')[1].split('-')
                date = QDate(date[0], date[1], date[2])   # YYYY, MM, DD
                optional = readfile_setting.readline().rstrip('\n').split('=')[1]
                self.set_line_texts(location_save=location_load, analyst=analyst, date=date, optional=optional)
            else:
                print('No information file in the folder.')
                return
        else:
            QMessageBox.information(self, "Error", "Select a folder.")
            return

    def click_btn_exit(self):
        sys.exit()

    def click_btn_ok(self):
        return


def terminal():
    readfile_setting = open(os.getcwd()+'/Settings/Settings.ini', mode='r')  # File I/O Error 가능성
    location_save = readfile_setting.readline().rstrip('\n').split('=')[1]
    analyst = readfile_setting.readline().rstrip('\n').split('=')[1]
    readfile_setting.close()
    os.chdir(location_save)
    print("#################")
    print("Cherry Tomatoe")
    print("#################")
    while True:
        print('1.New\n2.Load\n3.Exit')
        print('Choose 1, 2 or 3 : ')
        choice = input()
        if choice is '1':
            thisterm = DataDNAIdentification()
            thisterm.analyst = analyst
            os.mkdir(thisterm.get_dir())
            os.chdir(location_save+thisterm.get_dir())
            thisterm.export_info()
            break
        elif choice is '2':
            print("Enter the folder name : ")
            location_load = input()
            print('/%s/%s' % (location_load, 'info.ini'))
            if os.path.isdir(location_load) and os.path.exists('%s/%s' % (location_load, 'info.ini')):
                os.chdir('%s%s' % (location_save, location_load))
                thisterm = DataDNAIdentification()
                thisterm.import_info()
                break
            else:
                print('No folder or No info')
                continue
        elif choice == '3':
            print('Program closed')
            sys.exit()
        else:
            print('Choose 1, 2 or 3')


def except_hook(cls, exception, traceback): # for PyQt5.5 debugging
    sys.__excepthook__(cls, exception, traceback)


if __name__ == "__main__":
        app = QApplication(sys.argv)
        entry_GUI = EntryForm()
        entry_GUI.setFixedSize(entry_GUI.size())  # 창 크기 고정
        entry_GUI.show()
        sys.excepthook = except_hook    # for PyQt5.5 debugging
        app.exec_()

