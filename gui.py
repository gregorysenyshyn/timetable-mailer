import io
import os
import sys

from PyPDF2 import PdfFileReader

from PyQt5.QtCore import QThread
from PyQt5.QtCore import QObject
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QDialog
from PyQt5.QtWidgets import QCheckBox
from PyQt5.QtWidgets import QLineEdit
from PyQt5.QtWidgets import QTextEdit
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import QVBoxLayout
from PyQt5.QtWidgets import QHBoxLayout
from PyQt5.QtWidgets import QProgressBar
from PyQt5.QtWidgets import QApplication

import ttm


class Communicate(QObject):
    mailed = pyqtSignal(['QString'])
    finished = pyqtSignal()


class MainWindow(QMainWindow):

    communicate = Communicate()

    def __init__(self):
        QMainWindow.__init__(self)
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Email Maplewood Timetables to UGCloud')
        self.setMinimumWidth(800)
        main_widget = QWidget()
        self.layout = QVBoxLayout()
        main_widget.setLayout(self.layout)

        self.file_row = QHBoxLayout()
        file_input_label = QLabel('Timetable File: ')
        self.file_input = QLineEdit()
        file_browser_button = QPushButton('Choose File...')
        file_browser_button.clicked.connect(lambda: self.open_file_browser(
                                            self.file_input))
        self.file_row.addWidget(file_input_label)
        self.file_row.addWidget(self.file_input)
        self.file_row.addWidget(file_browser_button)
        self.layout.addLayout(self.file_row)

        self.mail_row = QHBoxLayout()
        email_input_label = QLabel('Your UGCloud Email: ')
        self.email_input = QLineEdit()
        self.email_input.setText('jdoe@ugcloud.ca')
        self.mail_row.addWidget(email_input_label)
        self.mail_row.addWidget(self.email_input)
        self.layout.addLayout(self.mail_row)

        self.password_row = QHBoxLayout()
        password_input_label = QLabel('Your UGCloud Password: ')
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_row.addWidget(password_input_label)
        self.password_row.addWidget(self.password_input)
        self.layout.addLayout(self.password_row)

        self.subject_row = QHBoxLayout()
        subject_input_label = QLabel('Subject for Emails: ')
        self.subject_input = QLineEdit()
        self.subject_input.setText('Enjoy your new timetable!')
        self.subject_row.addWidget(subject_input_label)
        self.subject_row.addWidget(self.subject_input)
        self.layout.addLayout(self.subject_row)

        self.message_label_row = QHBoxLayout()
        self.message_input_label = QLabel('Message')
        self.message_label_row.addWidget(self.message_input_label)
        self.layout.addLayout(self.message_label_row)

        self.message_input_row = QHBoxLayout()
        self.message_input = QTextEdit()
        self.message_input.setText(' '.join(['Please find your timetable',
                                             'attached to this email']))
        self.message_input_row.addWidget(self.message_input)
        self.layout.addLayout(self.message_input_row)

        self.go_row = QHBoxLayout()
        go_button = QPushButton('Send Timetables')
        go_button.clicked.connect(self.mail_timetables)
        self.go_row.addWidget(go_button)
        self.layout.addLayout(self.go_row)

        self.setCentralWidget(main_widget)

    def open_file_browser(self, input_to_change):
        self.file_dialog = QFileDialog()
        file_name = self.file_dialog.getOpenFileName(self,
                                                     'Choose Timetables',
                                                     os.path.expanduser('~'),
                                                     'PDF Files (*.pdf)')
        if file_name:
            input_to_change.setText(file_name[0])

    def init_output_display(self, pdf_page_count):
        output_window = QDialog()
        output_window.setWindowTitle('Mail Progress')
        layout = QVBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(pdf_page_count)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.output_text = QTextEdit()
        self.output_text.setMinimumHeight(300)
        self.output_text.setReadOnly(True)
        layout.addWidget(self.output_text)

        output_window.setLayout(layout)
        self.layout.addWidget(output_window)

    def update_output(self, address):
        self.output_text.insertPlainText('Sent email to {0}\n'.format(address))
        self.progress_bar.setValue(self.progress_bar.value() + 1)

    def finished(self):
        self.output_text.insertPlainText(' '.join(['Finished sending all',
                                                   'emails successfully\n']))

    def mail_timetables(self):

        pdf_file = self.file_input.text()
        email_user = self.email_input.text()
        email_pass = self.password_input.text()
        subject = self.subject_input.text()
        body = self.message_input.toPlainText()
        with open(pdf_file, 'rb') as f:
            pdf_io = io.BytesIO(f.read())
            reader = PdfFileReader(pdf_io)
            self.init_output_display(reader.numPages)
            self.get_thread = MailTimetablesThread(reader,
                                                   email_user,
                                                   email_pass,
                                                   subject,
                                                   self.communicate,
                                                   body)
            self.communicate.mailed.connect(self.update_output)
            self.get_thread.finished.connect(self.finished)
            self.output_text.insertPlainText(
                'Mailing {0} timetables\n'.format(reader.numPages))
            self.get_thread.start()


class MailTimetablesThread(QThread):

    def __init__(self, reader, email_user,
                 email_pass, subject, communicate, body):
        QThread.__init__(self)
        self.reader = reader
        self.email_user = email_user
        self.email_pass = email_pass
        self.subject = subject
        self.communicate = communicate
        self.body = body

    def run(self):
        oen_re = ttm.get_oen_re()
        smtp = ttm.get_smtp(self.email_user, self.email_pass)
        for page in self.reader.pages:
            msg = ttm.get_message(page,
                                  self.email_user,
                                  self.subject,
                                  oen_re,
                                  self.body)
            smtp.send_message(msg)

            self.communicate.mailed.emit(msg['To'])
            self.sleep(2)
        smtp.close()


if __name__ == '__main__':
    qt_app = QApplication(sys.argv)
    qt_app.setStyleSheet('QWidget { font-size:18px;}')
    app = MainWindow()

    msg_box = QMessageBox()
    msg_box.setText(" ".join(["Remember to use the students' 'given names'",
                              "when generating lists in Maplewood"]))
    msg_box.exec_()

    app.show()
    qt_app.exec_()
