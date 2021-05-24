import configparser
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMessageBox
from gui_test.log_in_window import StartWindow
from gui_test.main_window import Ui_MainWindow
from test import log_in, log_out, add_word, add_keyword_section, delete_word, delete_section, get_emails, \
    get_all_emails, find_urgent_emails, find_keywords_emails, get_attachment, check_labs, \
    check_course_projects, name_config
import sys


class start_window(QtWidgets.QMainWindow):
    service = 0
    def __init__(self):
        super(start_window, self).__init__()
        self.start = StartWindow()
        self.iniUI()

    def iniUI(self):
        self.start.setupUi(self)
        log_in_button = self.start.pushButton
        log_in_button.clicked.connect(self.logging)

    def logging(self):
        self.service = log_in()
        self.start.label_2.setText("Success!\nNow you can close this window")


class main_window(QtWidgets.QMainWindow):
    service = 0
    log_in_flag = 0

    def __init__(self):
        super(main_window, self).__init__()
        self.ui = Ui_MainWindow()
        # self.start = start_window()
        # self.start.show()
        self.iniUI()

    def iniUI(self):
        self.ui.setupUi(self)
        self.ui.buttonLogIn.clicked.connect(self.logIn)
        self.ui.logOut.clicked.connect(self.logOut)
        self.ui.addWord.clicked.connect(self.addWord)
        self.ui.addSection.clicked.connect(self.addSection)
        self.ui.getAllEmails.clicked.connect(self.printAllEmails)
        self.ui.getUrgentEmails.clicked.connect(self.printUrgentEmails)
        self.ui.getKeywordEmails.clicked.connect(self.printKeywordEmails)
        self.ui.downloadAttchment.clicked.connect(self.downloadAttachment)
        self.ui.showSections.clicked.connect(self.showSections)
        self.ui.checkLabs.clicked.connect(self.checkLabs)
        self.ui.checkProjects.clicked.connect(self.checkCourseProjects)

    def checkForService(self):
        if self.service == 0:
            self.errorMessage("You need to log in for further work\nPlease, log in on the first tab")
            return 1
        return 0

    def getFiletypeEmail(self):
        file_type = 0
        if self.ui.allEmailsFile.currentText() == ".json":
            file_type = 1
        elif self.ui.allEmailsFile.currentText() == ".txt":
            file_type = 0
        return file_type

    def getParametrsEmails(self):
        query = ''
        if self.ui.editAfter.text() != '':
            query += 'after:' + self.ui.editAfter.text()
        if self.ui.editBefore.text() != '':
            query += ' before:' + self.ui.editAfter.text()
        if self.ui.editFrom.text() != '':
            query += ' from:' + self.ui.editAfter.text()
        if self.ui.editUnread.text() != '':
            query += ' is:' + self.ui.editAfter.text()
        return query

    def logIn(self):
        self.service = log_in()
        self.log_in_flag = 1
        self.ui.label_30.setText("Success!\nNow you can go on\nanother tab")

    def logOut(self):
        if self.service == 0:
            self.errorMessage("You haven't log in yet")
            return
        if self.log_in_flag != 0:
            log_out()
            self.log_in_flag = 0

    def addWord(self):
        word = self.ui.editWord1.text()
        if word == '':
            self.errorMessage("You need to write keywords")
            return
        section = self.ui.editSection1.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        words = word.split(', ')
        for i in words:
            add_word(section, i)
        self.ui.editWord1.setText('')
        self.ui.editSection1.setText('')

    def addSection(self):
        words = self.ui.editWord2.text()
        if words == '':
            self.errorMessage("You need to write keywords")
            return
        section = self.ui.editSection2.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        add_keyword_section(section, words.split(', '))
        self.ui.editWord2.setText('')
        self.ui.editSection2.setText('')

    def deleteWord(self):
        word = self.ui.editWord1_2.text()
        if word == '':
            self.errorMessage("You need to write keywords")
            return
        section = self.ui.editSection1_2.text()
        words = word.split(', ')
        for i in words:
            delete_word(section, i)
        self.ui.editWord1_2.setText('')
        self.ui.editSection1_2.setText('')

    def deleteSection(self):
        section = self.ui.editSection2_2.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        delete_section(section)
        self.ui.editSection2_2.setText('')

    def showSections(self):
        self.ui.listSections.clear()
        self.ui.listWords.clear()
        config = configparser.ConfigParser()
        config.read(name_config)
        self.ui.listSections.addItems(config.sections())
        self.ui.listSections.itemClicked.connect(self.sectionChoice)

    def sectionChoice(self, item):
        self.ui.listWords.clear()
        config = configparser.ConfigParser()
        config.read(name_config)
        result = list()
        words = config.items(item.text())
        for word in words:
            result.append(word[1])
        self.ui.listWords.addItems(result)

    def printAllEmails(self):
        if self.checkForService() == 1:
            return
        file_type = self.getFiletypeEmail()
        query = self.getParametrsEmails()
        path = self.ui.editPath.text()
        get_all_emails(self.service, query, file_type, path)

    def printUrgentEmails(self):
        if self.checkForService() == 1:
            return
        file_type = self.getFiletypeEmail()
        query = self.getParametrsEmails()
        path = self.ui.editPath.text()
        find_urgent_emails(self.service, query, file_type, path)

    def printKeywordEmails(self):
        if self.checkForService() == 1:
            return
        file_type = self.getFiletypeEmail()
        query = self.getParametrsEmails()
        section = self.ui.keywordName.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        path = self.ui.editPath.text()
        filename = self.ui.keywordFileName.text()
        if filename == '':
            filename = str(section)
        find_keywords_emails(self.service, query, section, filename, file_type, path)

    def downloadAttachment(self):
        if self.checkForService() == 1:
            return
        messageID = self.ui.messageID.text()
        attachmentID = self.ui.attachmentID.text()
        path = self.ui.pathForAttachment.text()
        filename = self.ui.attachmentName.text()
        get_attachment(self.service, messageID, attachmentID, path, filename)

    def checkLabs(self):
        if self.checkForService() == 1:
            return
        num = self.ui.editNumberLabs.text()
        filename = self.ui.editLabsFilename.text()
        path = self.ui.editPathCheck.text()
        query = self.getParametrsEmails()
        check_labs(get_emails(self.service, query), num, filename, path)

    def checkCourseProjects(self):
        if self.checkForService() == 1:
            return
        num = self.ui.editGroupNumber.text()
        filename = self.ui.editProjectsFilename.text()
        filetype = self.ui.editProjectsType.text()
        path = self.ui.editPathCheck.text()
        query = self.getParametrsEmails()
        check_course_projects(get_emails(self.service, query), filename, num, filetype, path)

    def errorMessage(self, text):   # окно ошибки
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Critical)
        msgBox.setText(text)
        msgBox.setWindowTitle("Error")
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec()


def show_start():
    window = start_window()
    window.show()


app = QtWidgets.QApplication([])
application = main_window()
application.show()
# window = start_window()
# window.show()

sys.exit(app.exec())
