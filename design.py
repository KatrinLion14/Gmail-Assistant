from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMessageBox
from main_window import Ui_MainWindow
from functions import *
import sys


class main_window(QtWidgets.QMainWindow):
    service = 0
    emails = []

    def __init__(self):
        super(main_window, self).__init__()
        self.ui = Ui_MainWindow()
        self.iniUI()

    def iniUI(self):
        self.setWindowIcon(QIcon("icon.png"))
        self.ui.setupUi(self)
        self.ui.buttonLogIn.clicked.connect(self.logIn)
        self.ui.logOut.clicked.connect(self.logOut)
        self.ui.addWord.clicked.connect(self.addWord)
        self.ui.addSection.clicked.connect(self.addSection)
        self.ui.deleteWord.clicked.connect(self.deleteWord)
        self.ui.deleteSection.clicked.connect(self.deleteSection)
        self.ui.getAllEmails.clicked.connect(self.printAllEmails)
        self.ui.getUrgentEmails.clicked.connect(self.printUrgentEmails)
        self.ui.getKeywordEmails.clicked.connect(self.printKeywordEmails)
        self.ui.downloadAttchment.clicked.connect(self.downloadAttachment)
        self.ui.showSections.clicked.connect(self.showSections)
        self.ui.checkLabs.clicked.connect(self.checkLabs)
        self.ui.checkProjects.clicked.connect(self.checkCourseProjects)
        self.ui.getAllEmails_2.clicked.connect(self.showEmails)
        self.ui.getKeywordEmails_2.clicked.connect(self.showKeywordEmails)
        self.ui.downloadAttchment_2.clicked.connect(self.downloadAttachments)

    def checkForService(self):
        if self.service == 0:
            self.errorMessage("You need to log in for further work\nPlease, log in on the first tab")
            return 1
        return 0

    def checkIfAll(self, query):
        if query == '':
            warning = self.warningMessage("Are you sure you want to\ncheck all emails?\nIt may take some time")
            if warning == 0:
                return 1
        return 0

    def getParametrsEmails(self):
        query = ''
        if self.ui.editAfter.text() != '':
            query += 'after:' + self.ui.editAfter.text()
        if self.ui.editBefore.text() != '':
            query += ' before:' + self.ui.editBefore.text()
        if self.ui.editFrom.text() != '':
            query += ' from:' + self.ui.editFrom.text()
        if self.ui.editUnread.text() != '':
            query += ' is:' + self.ui.editUnread.text()
        return query

    def getParametrsCheck(self):
        query = ''
        if self.ui.editAfterCheck.text() != '':
            query += 'after:' + self.ui.editAfterCheck.text()
        if self.ui.editBeforeCheck.text() != '':
            query += ' before:' + self.ui.editBeforeCheck.text()
        if self.ui.editFromCheck.text() != '':
            query += ' from:' + self.ui.editFromCheck.text()
        if self.ui.editUnreadCheck.text() != '':
            query += ' is:' + self.ui.editUnreadCheck.text()
        return query

    def getParametrsDownload(self):
        query = ''
        if self.ui.editAfter_2.text() != '':
            query += 'after:' + self.ui.editAfter_2.text()
        if self.ui.editBefore_2.text() != '':
            query += ' before:' + self.ui.editBefore_2.text()
        if self.ui.editFrom_2.text() != '':
            query += ' from:' + self.ui.editFrom_2.text()
        if self.ui.editUnread_2.text() != '':
            query += ' is:' + self.ui.editUnread_2.text()
        return query

    def logIn(self):
        self.service = log_in()
        self.ui.label_30.setText("Success!\nNow you can go on\nanother tab")
        self.ui.label_31.setText("")

    def logOut(self):
        if self.service == 0:
            self.errorMessage("You haven't log in yet")
            return
        else:
            log_out()
            self.ui.label_31.setText("You successfully logged out!")
            self.ui.label_30.setText("You haven't logged in yet\nYou cannot process further")

    def addWord(self):
        self.ui.success1.setText("")
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
            try:
                add_word(section, i)
            except:
                self.errorMessage("Something gone wrong\nPlease, check the input and try again")
                return
        self.ui.editWord1.setText('')
        self.ui.editSection1.setText('')
        self.ui.success1.setText("Success!")

    def addSection(self):
        self.ui.success2.setText("")
        words = self.ui.editWord2.text()
        if words == '':
            self.errorMessage("You need to write keywords")
            return
        section = self.ui.editSection2.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        try:
            add_keyword_section(section, words.split(', '))
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")
            return
        self.ui.editWord2.setText('')
        self.ui.editSection2.setText('')
        self.ui.success2.setText("Success!")

    def deleteWord(self):
        self.ui.success3.setText("")
        word = self.ui.editWord1_2.text()
        if word == '':
            self.errorMessage("You need to write keywords")
            return
        section = self.ui.editSection1_2.text()
        words = word.split(', ')
        for i in words:
            try:
                delete_word(section, i)
            except:
                self.errorMessage("Something gone wrong\nPlease, check the input and try again")
                return
        self.ui.editWord1_2.setText('')
        self.ui.editSection1_2.setText('')
        self.ui.success3.setText("Success!")

    def deleteSection(self):
        self.ui.success4.setText("")
        section = self.ui.editSection2_2.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        try:
            delete_section(section)
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")
            return
        self.ui.editSection2_2.setText('')
        self.ui.success4.setText("Success!")

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
        self.ui.success5.setText("")
        if self.checkForService() == 1:
            return
        file_type = 0
        if self.ui.allEmailsFile.currentText() == ".json":
            file_type = 1
        elif self.ui.allEmailsFile.currentText() == ".txt":
            file_type = 0
        query = self.getParametrsEmails()
        path = self.ui.editPath.text()
        if self.checkIfAll(query) == 1:
            return
        filename = self.ui.allEmailsFileName.text()
        if filename == '':
            filename = "all_emails"
        try:
            get_all_emails(self.service, query, file_type, path, filename)
            self.ui.success5.setText("Success!")
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")

    def printUrgentEmails(self):
        self.ui.success6.setText("")
        if self.checkForService() == 1:
            return
        file_type = 0
        if self.ui.urgentEmailsFile.currentText() == ".json":
            file_type = 1
        elif self.ui.urgentEmailsFile.currentText() == ".txt":
            file_type = 0
        query = self.getParametrsEmails()
        if self.checkIfAll(query) == 1:
            return
        path = self.ui.editPath.text()
        filename = self.ui.urgentFileName.text()
        if filename == '':
            filename = "urgent"
        try:
            find_urgent_emails(self.service, query, file_type, path, filename)
            self.ui.success6.setText("Success!")
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")

    def printKeywordEmails(self):
        self.ui.success7.setText("")
        if self.checkForService() == 1:
            return
        file_type = 0
        if self.ui.keywordEmailsFile.currentText() == ".json":
            file_type = 1
        elif self.ui.keywordEmailsFile.currentText() == ".txt":
            file_type = 0
        query = self.getParametrsEmails()
        if self.checkIfAll(query) == 1:
            return
        section = self.ui.keywordName.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        path = self.ui.editPath.text()
        filename = self.ui.keywordFileName.text()
        if filename == '':
            filename = str(section)
        try:
            find_keywords_emails(self.service, query, section, filename, file_type, path)
            self.ui.success7.setText("Success!")
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")

    def downloadAttachment(self):
        self.ui.success8.setText("")
        if self.checkForService() == 1:
            return
        messageID = self.ui.messageID.text()
        if messageID == '':
            self.errorMessage("You need to write email's ID")
            return
        attachmentID = self.ui.attachmentID.text()
        path = self.ui.pathForAttachment.text()
        filename = self.ui.attachmentName.text()
        if filename == '' and attachmentID != '':
            self.errorMessage("You need to choose name for file")
            return
        try:
            get_attachment(self.service, messageID, attachmentID, path, filename)
            self.ui.success8.setText("Success!")
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")

    def showEmails(self):
        self.ui.listWidget.clear()
        self.ui.listWidget_2.clear()
        if self.checkForService() == 1:
            return
        self.ui.listWidget.clear()
        query = self.getParametrsDownload()
        if self.checkIfAll(query) == 1:
            return
        try:
            self.emails = get_emails_for_show(self.service, query)
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")
        if len(self.emails) == 0:
            self.errorMessage("Cannot find any emails")
            return
        for key in self.emails:
            item = QtWidgets.QListWidgetItem()
            item.setText(self.emails[key]['from'] + "(" + self.emails[key]['id'] + ")")
            item.setCheckState(QtCore.Qt.Unchecked)
            self.ui.listWidget.addItem(item)
        self.ui.listWidget.itemClicked.connect(self.showText)

    def showKeywordEmails(self):
        self.ui.listWidget.clear()
        self.ui.listWidget_2.clear()
        if self.checkForService() == 1:
            return
        self.ui.listWidget.clear()
        query = self.getParametrsDownload()
        if self.checkIfAll(query) == 1:
            return
        section = self.ui.keywordName_2.text()
        if section == '':
            self.errorMessage("You need to write keywords section")
            return
        try:
            self.emails = find_keywords_emails_for_show(self.service, query, section)
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")
        if len(self.emails) == 0:
            self.errorMessage("No emails with such words")
            return
        for key in self.emails:
            item = QtWidgets.QListWidgetItem()
            item.setText(self.emails[key]['from'] + "(" + self.emails[key]['id'] + ")")
            item.setCheckState(QtCore.Qt.Unchecked)
            self.ui.listWidget.addItem(item)
        self.ui.listWidget.itemClicked.connect(self.showText)

    def showText(self, item):
        self.ui.listWidget_2.clear()
        id = item.text().split('(')[1]
        mail = self.emails[id[:-1]]
        result = ""
        for key in mail:
            result += unicode(key) + ' : ' + unicode(mail[key]) + '\n'
        item1 = QtWidgets.QListWidgetItem()
        item1.setText(result)
        self.ui.listWidget_2.addItem(item1)

    def downloadAttachments(self):
        if self.checkForService() == 1:
            return
        path = self.ui.pathForAttachment_3.text()
        filetype = self.ui.editFileType_2.text()
        errors = []
        result = []
        for i in range(self.ui.listWidget.count()):
            if self.ui.listWidget.item(i).checkState():
                id = self.ui.listWidget.item(i).text().split('(')[1]
                mail = self.emails[id[:-1]]
                if filetype == '':
                    try:
                        get_attachment(self.service, id[:-1], '', path, '')
                        result.append(id[:-1])
                    except:
                        errors.append(id[:-1])
                else:
                    downloaded = False
                    for j in range(len(mail['attachment_names'])):
                        if mail['attachment_names'][j].find(filetype) > -1:
                            downloaded = True
                            try:
                                get_attachment(self.service, id[:-1], mail['attachment_ids'][j], path, mail['attachment_names'][j])
                            except:
                                errors.append(id[:-1])
                    if downloaded:
                            result.append(id[:-1])
        string = "Downloaded from: "
        for word in result:
            string = string + word + "; "
        string += "\nErrors in: "
        for word in errors:
            string = string + word + "; "
        self.infoMessage(string)


    def checkLabs(self):
        self.ui.success9.setText("")
        if self.checkForService() == 1:
            return
        num = self.ui.editNumberLabs.text()
        filename = self.ui.editLabsFilename.text()
        path = self.ui.editPathCheck.text()
        query = self.getParametrsCheck()
        if self.checkIfAll(query) == 1:
            return
        try:
            check_labs(get_emails(self.service, query), num, filename, path)
            self.ui.success9.setText("Success!")
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")

    def checkCourseProjects(self):
        self.ui.success10.setText("")
        if self.checkForService() == 1:
            return
        num = self.ui.editGroupNumber.text()
        filename = self.ui.editProjectsFilename.text()
        filetype = self.ui.editProjectsType.text()
        path = self.ui.editPathCheck.text()
        query = self.getParametrsCheck()
        if self.checkIfAll(query) == 1:
            return
        try:
            check_course_projects(get_emails(self.service, query), filename, num, filetype, path)
            self.ui.success10.setText("Success!")
        except:
            self.errorMessage("Something gone wrong\nPlease, check the input and try again")

    def errorMessage(self, text):   # окно ошибки
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Critical)
        msgBox.setWindowIcon(QIcon("icon.png"))
        msgBox.setText(text)
        msgBox.setWindowTitle("Error")
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec()

    def warningMessage(self, text):   # окно предупреждения
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Warning)
        msgBox.setWindowIcon(QIcon("icon.png"))
        msgBox.setText(text)
        msgBox.setWindowTitle("Warning")
        msgBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        if msgBox.exec() == QMessageBox.Yes:
            return 1
        else:
            return 0

    def infoMessage(self, text):   # окно информации
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information)
        msgBox.setWindowIcon(QIcon("icon.png"))
        msgBox.setText(text)
        msgBox.setWindowTitle("Info")
        msgBox.setStandardButtons(QMessageBox.Ok)
        msgBox.exec()

app = QtWidgets.QApplication([])
application = main_window()
application.show()

sys.exit(app.exec())
