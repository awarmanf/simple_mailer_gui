from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

from qfluentwidgets import EditableComboBox
from qfluentwidgets import PushButton
from qfluentwidgets import PrimaryPushButton
from qfluentwidgets import RadioButton
from qfluentwidgets import HorizontalSeparator
from qfluentwidgets import BodyLabel
from qfluentwidgets import TitleLabel
from qfluentwidgets import LineEdit
from qfluentwidgets import PasswordLineEdit
from qfluentwidgets import PlainTextEdit
from qfluentwidgets import RoundMenu, Action, MenuAnimationType
from qfluentwidgets import Dialog
from qfluentwidgets import FluentIcon as FIF

from qframelesswindow import FramelessWindow

import base64, json, os

import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
from email import encoders

from images_data import IconApp


class WorkerSendMail(QThread):
    error = pyqtSignal(dict)
    success = pyqtSignal(dict)
    
    def __init__(self, name, account, password, smtp_host, security_type, recipients, subject, body, file_attachment, parent=None):
        super().__init__(parent=parent)
        self.name = name
        self.account = account
        self.password = password
        self.smtp_host = smtp_host
        self.security_type = security_type
        self.recipients = recipients
        self.subject = subject
        self.body = body
        self.file_attachment = file_attachment

    def run(self):
        try:
            server = self.check_connection(self.smtp_host)
            if server:
                server.login(self.account, self.password)
                self.send_mail(server)
        except Exception as e:
            self.error.emit({'title': 'Error', 'content': f'Error: {e}'})

    def check_connection(self, host):
        try:
            if self.security_type == 'plain':
                server = smtplib.SMTP(host, port=25, timeout=10)
            elif self.security_type == 'SSL':
                server = smtplib.SMTP_SSL(host, port=465, timeout=10)
            elif self.security_type == 'TLS':
                server = smtplib.SMTP(host, port=587, timeout=10)
                server.starttls()
            else:
                self.error.emit({'title': 'Connection Error', 'content': 'Invalid security type.'})
                return None
            return server
        except Exception as e:
            self.error.emit({'title': 'Connection Error', 'content': f'Error: {e}'})

    def send_mail(self, server):
        try:
            msg = email.mime.multipart.MIMEMultipart()
            msg['From'] = self.account
            msg['To'] = self.recipients
            msg['Subject'] = self.subject

            msg.attach(email.mime.text.MIMEText(self.body, 'plain'))

            if self.file_attachment:
                with open(self.file_attachment, 'rb') as f:
                    part = email.mime.base.MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(self.file_attachment)}')
                    msg.attach(part)

            # Send email
            server.sendmail(self.account, self.recipients, msg.as_string())
            server.quit()
            self.success.emit({'title': 'Success', 'content': 'Mail sent successfully'})
        except Exception as e:
            self.error.emit({'title': 'Send Mail Error', 'content': f'Error: {e}'})


class SimpleMailer(FramelessWindow):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.version = '1.3'
        self.thisrelease = '2024'
        self.config = (lambda: (json.load(open('config.json', 'r', encoding='utf-8')) if os.path.exists('config.json') else {'full_name': '', 'account_data': {'1': {'account_mail': '', 'account_password': ''}}, 'SMTP_host': '', 'security_type': ''}))()
        if not os.path.exists('config.json'):
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
        
        self.GUI()
    
    def GUI(self):
        self.setObjectName(u"SimpleMailerApp")
        self.setWindowIcon(QIcon(QPixmap.fromImage(QImage.fromData(base64.b64decode(IconApp)))))
        self.resize(476, 760)
        self.gridLayout = QGridLayout(self)
        self.gridLayout.setObjectName(u"gridLayout")
        self.gridLayout.setContentsMargins(-1, 35, -1, -1)
        self.input_password = PasswordLineEdit(self)
        self.input_password.setObjectName(u"input_password")

        self.gridLayout.addWidget(self.input_password, 4, 1, 1, 3)

        self.label_recipients = BodyLabel(self)
        self.label_recipients.setObjectName(u"label_recipients")
        self.label_recipients.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_recipients, 9, 0, 1, 1)

        self.input_body = PlainTextEdit(self)
        self.input_body.setObjectName(u"input_body")

        self.gridLayout.addWidget(self.input_body, 12, 0, 1, 4)

        self.input_SMTP_host = LineEdit(self)
        self.input_SMTP_host.setObjectName(u"input_SMTP_host")

        self.gridLayout.addWidget(self.input_SMTP_host, 5, 1, 1, 3)

        self.label_file_attachment = BodyLabel(self)
        self.label_file_attachment.setObjectName(u"label_file_attachment")
        self.label_file_attachment.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_file_attachment, 10, 0, 1, 1)

        self.input_recipients = LineEdit(self)
        self.input_recipients.setObjectName(u"input_recipients")

        self.gridLayout.addWidget(self.input_recipients, 9, 1, 1, 3)

        self.HorizontalSeparator = HorizontalSeparator(self)
        self.HorizontalSeparator.setObjectName(u"HorizontalSeparator")

        self.gridLayout.addWidget(self.HorizontalSeparator, 1, 0, 1, 4)

        self.button_close = PushButton(self)
        self.button_close.setObjectName(u"button_close")
        self.button_close.setMaximumSize(QSize(200, 16777215))

        self.gridLayout.addWidget(self.button_close, 13, 3, 1, 1)

        self.label_full_name = BodyLabel(self)
        self.label_full_name.setObjectName(u"label_full_name")
        self.label_full_name.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_full_name, 2, 0, 1, 1)

        self.input_file_attachment = LineEdit(self)
        self.input_file_attachment.setObjectName(u"input_file_attachment")

        self.gridLayout.addWidget(self.input_file_attachment, 10, 1, 1, 2)

        self.label_subject = BodyLabel(self)
        self.label_subject.setObjectName(u"label_subject")
        self.label_subject.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_subject, 8, 0, 1, 1)

        self.input_full_name = LineEdit(self)
        self.input_full_name.setObjectName(u"input_full_name")

        self.gridLayout.addWidget(self.input_full_name, 2, 1, 1, 3)

        self.input_subject = LineEdit(self)
        self.input_subject.setObjectName(u"input_subject")

        self.gridLayout.addWidget(self.input_subject, 8, 1, 1, 3)

        self.button_select_file_attachment = PushButton(self)
        self.button_select_file_attachment.setObjectName(u"button_select_file_attachment")
        self.button_select_file_attachment.setMaximumSize(QSize(250, 16777215))

        self.gridLayout.addWidget(self.button_select_file_attachment, 10, 3, 1, 1)

        self.label_SMTP_host = BodyLabel(self)
        self.label_SMTP_host.setObjectName(u"label_SMTP_host")
        self.label_SMTP_host.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_SMTP_host, 5, 0, 1, 1)

        self.button_submit = PrimaryPushButton(self)
        self.button_submit.setObjectName(u"button_submit")
        self.button_submit.setMaximumSize(QSize(200, 16777215))

        self.gridLayout.addWidget(self.button_submit, 13, 2, 1, 1)

        self.label_password = BodyLabel(self)
        self.label_password.setObjectName(u"label_password")
        self.label_password.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_password, 4, 0, 1, 1)

        self.label_security_type = BodyLabel(self)
        self.label_security_type.setObjectName(u"label_security_type")
        self.label_security_type.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_security_type, 7, 0, 1, 1)

        self.input_account = EditableComboBox(self)
        self.input_account.setObjectName(u"input_account")

        self.gridLayout.addWidget(self.input_account, 3, 1, 1, 3)

        self.label_account = BodyLabel(self)
        self.label_account.setObjectName(u"label_account")
        self.label_account.setMaximumSize(QSize(120, 16777215))

        self.gridLayout.addWidget(self.label_account, 3, 0, 1, 1)

        self.title_simple_mailer = TitleLabel(self)
        self.title_simple_mailer.setObjectName(u"title_simple_mailer")

        self.gridLayout.addWidget(self.title_simple_mailer, 0, 0, 1, 4, Qt.AlignHCenter|Qt.AlignVCenter)

        self.frame = QFrame(self)
        self.frame.setObjectName(u"frame")
        self.frame.setFrameShape(QFrame.StyledPanel)
        self.frame.setFrameShadow(QFrame.Raised)
        self.gridLayout_2 = QGridLayout(self.frame)
        self.gridLayout_2.setObjectName(u"gridLayout_2")
        self.option_security_type_plain = RadioButton(self.frame)
        self.option_security_type_plain.setObjectName(u"option_security_type_plain")

        self.gridLayout_2.addWidget(self.option_security_type_plain, 0, 0, 1, 1, Qt.AlignHCenter)

        self.option_security_type_SSL = RadioButton(self.frame)
        self.option_security_type_SSL.setObjectName(u"option_security_type_SSL")

        self.gridLayout_2.addWidget(self.option_security_type_SSL, 0, 2, 1, 1, Qt.AlignHCenter)

        self.option_security_type_TLS = RadioButton(self.frame)
        self.option_security_type_TLS.setObjectName(u"option_security_type_TLS")

        self.gridLayout_2.addWidget(self.option_security_type_TLS, 0, 1, 1, 1, Qt.AlignHCenter)


        self.gridLayout.addWidget(self.frame, 7, 1, 1, 3)

        self.label_body = BodyLabel(self)
        self.label_body.setObjectName(u"label_body")
        self.label_body.setMaximumSize(QSize(16777215, 16777215))

        self.gridLayout.addWidget(self.label_body, 11, 0, 1, 4)

        self.setWindowTitle(QCoreApplication.translate("SimpleMailer", u"Simple Mailer", None))
        self.label_recipients.setText(QCoreApplication.translate("SimpleMailer", u"Recipients", None))
        self.label_file_attachment.setText(QCoreApplication.translate("SimpleMailer", u"File Attachment", None))
        self.button_close.setText(QCoreApplication.translate("SimpleMailer", u"Close", None))
        self.label_full_name.setText(QCoreApplication.translate("SimpleMailer", u"Full Name", None))
        self.label_subject.setText(QCoreApplication.translate("SimpleMailer", u"Subject", None))
        self.button_select_file_attachment.setText(QCoreApplication.translate("SimpleMailer", u"Select File", None))
        self.label_SMTP_host.setText(QCoreApplication.translate("SimpleMailer", u"SMTP Host", None))
        self.button_submit.setText(QCoreApplication.translate("SimpleMailer", u"Submit", None))
        self.label_password.setText(QCoreApplication.translate("SimpleMailer", u"Password", None))
        self.label_security_type.setText(QCoreApplication.translate("SimpleMailer", u"Security Type", None))
        self.label_account.setText(QCoreApplication.translate("SimpleMailer", u"Account", None))
        self.title_simple_mailer.setText(QCoreApplication.translate("SimpleMailer", u"Simple Mailer", None))
        self.option_security_type_plain.setText(QCoreApplication.translate("SimpleMailer", u"Plain", None))
        self.option_security_type_SSL.setText(QCoreApplication.translate("SimpleMailer", u"SSL", None))
        self.option_security_type_TLS.setText(QCoreApplication.translate("SimpleMailer", u"TLS", None))
        self.label_body.setText(QCoreApplication.translate("SimpleMailer", u"Body:", None))
        
        QMetaObject.connectSlotsByName(self)

        self.input_file_attachment.setReadOnly(True)
        self.LoadConfig()

        #text changed event
        self.input_full_name.textChanged.connect(self.SaveConfigFullname)
        self.input_password.textChanged.connect(self.SaveConfigPassword)
        self.input_SMTP_host.textChanged.connect(self.SaveConfigSMTPHost)
        self.option_security_type_plain.toggled.connect(self.SaveConfigSecurityType)
        self.option_security_type_SSL.toggled.connect(self.SaveConfigSecurityType)
        self.option_security_type_TLS.toggled.connect(self.SaveConfigSecurityType)

        #button clicked event
        self.button_select_file_attachment.clicked.connect(self.OpenFileAttachment)
        self.button_close.clicked.connect(self.CloseApp)
        
        #index changed event
        self.input_account.currentIndexChanged.connect(self.ChangePassword)

        self.button_submit.clicked.connect(self.startSendMail)
        
    def ChangePassword(self):
        self.input_password.setText(self.config['account_data'][str(self.input_account.currentIndex()+1)]['account_password'])

    def contextMenuEvent(self, e):
        menu = RoundMenu(parent=self)
        menu.addAction(Action(FIF.ADD, 'Open', shortcut='Ctrl+O'))
        menu.addAction(Action(FIF.INFO, 'About', shortcut='Ctrl+I'))
        menu.addSeparator()
        menu.addAction(Action(FIF.CLOSE, 'Exit', shortcut='Ctrl+Q'))

        menu.actions()[0].triggered.connect(self.OpenFileAttachment)
        menu.actions()[1].triggered.connect(self.AboutApp)
        menu.actions()[2].triggered.connect(self.CloseApp)

        menu.exec(e.globalPos(), aniType=MenuAnimationType.DROP_DOWN)

    def OpenFileAttachment(self):
        file = QFileDialog.getOpenFileName(self, 'Open File Attachment', os.getenv('HOME'), 'All Files (*.*)')
        self.input_file_attachment.setText(file[0])

    def AboutApp(self):
        title = 'About Simple Mailer'
        content = f"""A Simple Mailer in PyQt5 QFluentWidget version {self.version} released in {self.thisrelease}.\n\nMain Developer: Arief Yudhawarman, Github: @awarmanf\n\nContributor: Arif Maulana Azis, Github: @Arifmaulanaazis"""
        w = Dialog(title, content, self)
        w.cancelButton.hide()
        w.exec()
        
    def CloseApp(self):
        self.close()

    def SaveConfigFullname(self, e):
        self.config['full_name'] = e
        json.dump(self.config, open('config.json', 'w', encoding='utf-8'), indent=4)


    def SaveConfigPassword(self, e):
        self.config['account_data'][str(self.input_account.currentIndex()+1)]['account_password'] = e
        json.dump(self.config, open('config.json', 'w', encoding='utf-8'), indent=4)

    def SaveConfigSMTPHost(self, e):
        self.config['SMTP_host'] = e
        json.dump(self.config, open('config.json', 'w', encoding='utf-8'), indent=4)

    def SaveConfigSecurityType(self):
        if self.option_security_type_plain.isChecked():
            self.config['security_type'] = 'plain'
        elif self.option_security_type_SSL.isChecked():
            self.config['security_type'] = 'SSL'
        elif self.option_security_type_TLS.isChecked():
            self.config['security_type'] = 'TLS'
        json.dump(self.config, open('config.json', 'w', encoding='utf-8'), indent=4)


    def LoadConfig(self):
        self.input_full_name.setText(self.config['full_name'])
        
        for i in self.config['account_data']:
            self.input_account.addItem(self.config['account_data'][i]['account_mail'])
            self.input_password.setText(self.config['account_data'][i]['account_password'])


        self.input_SMTP_host.setText(self.config['SMTP_host'])
        if self.config['security_type'] == 'plain':
            self.option_security_type_plain.setChecked(True)
        elif self.config['security_type'] == 'SSL':
            self.option_security_type_SSL.setChecked(True)
        elif self.config['security_type'] == 'TLS':
            self.option_security_type_TLS.setChecked(True)


    def closeEvent(self, event):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.quit()  
            self.worker.wait()  


        self.SaveConfigFullname(self.input_full_name.text())
        self.SaveConfigPassword(self.input_password.text())
        self.SaveConfigSMTPHost(self.input_SMTP_host.text())
        self.SaveConfigSecurityType()

        event.accept()


    def show_error_window(self, error_data):
        title = error_data['title']
        content = error_data['content']
        w = Dialog(title, content, self)
        w.cancelButton.hide() 
        w.exec()

    def show_success_window(self, success_data):
        title = success_data['title']
        content = success_data['content']
        w = Dialog(title, content, self)
        w.cancelButton.hide()
        w.exec()



    def startSendMail(self):
        self.DisableAllInput()
        self.worker = WorkerSendMail(
            name=self.input_full_name.text(),
            account=self.input_account.currentText(),
            password=self.input_password.text(),
            smtp_host=self.input_SMTP_host.text(),
            security_type=self.config['security_type'],
            recipients=self.input_recipients.text(),
            subject=self.input_subject.text(),
            body=self.input_body.toPlainText(),
            file_attachment=self.input_file_attachment.text()
        )
        
        self.worker.error.connect(self.Error_Window)
        self.worker.success.connect(self.Success_Window)
        
        self.worker.start()

    def Error_Window(self, error_data):
        self.EnableAllInput()
        self.show_error_window(error_data)

    def Success_Window(self, success_data):
        self.EnableAllInput()
        self.show_success_window(success_data)



    def DisableAllInput(self):
        self.input_full_name.setDisabled(True)
        self.input_account.setDisabled(True)
        self.input_password.setDisabled(True)
        self.input_SMTP_host.setDisabled(True)
        self.option_security_type_plain.setDisabled(True)
        self.option_security_type_SSL.setDisabled(True)
        self.option_security_type_TLS.setDisabled(True)
        self.input_recipients.setDisabled(True)
        self.input_subject.setDisabled(True)
        self.input_body.setDisabled(True)
        self.input_file_attachment.setDisabled(True)
        self.button_select_file_attachment.setDisabled(True)
        self.button_submit.setDisabled(True)

    def EnableAllInput(self):
        self.input_full_name.setDisabled(False)
        self.input_account.setDisabled(False)
        self.input_password.setDisabled(False)
        self.input_SMTP_host.setDisabled(False)
        self.option_security_type_plain.setDisabled(False)
        self.option_security_type_SSL.setDisabled(False)
        self.option_security_type_TLS.setDisabled(False)
        self.input_recipients.setDisabled(False)
        self.input_subject.setDisabled(False)
        self.input_body.setDisabled(False)
        self.input_file_attachment.setDisabled(False)
        self.button_select_file_attachment.setDisabled(False)
        self.button_submit.setDisabled(False)


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = SimpleMailer()
    window.show()
    sys.exit(app.exec_())