import configparser
import os
import sys
from PyQt5 import QtWidgets
from PyQt5.QtCore import QRunnable, QStringListModel, QThreadPool
from PyQt5.QtWidgets import QMainWindow, QAbstractItemView, QApplication, QFileDialog, QMessageBox
from GUI import Ui_MainWindow
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import datetime

# 自定义多线程类Start
class Start(QRunnable):
    def __init__(self, func):
        super().__init__()
        self.func = func

    def run(self):
        self.func()

class LoginClass(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.retranslateUi(self)
        # 连接pushButton按钮的点击事件与自定义槽函数
        self.pushButton.clicked.connect(self.start_on_push_button_clicked)
        self.pushButton_2.clicked.connect(self.save_config)
        self.pushButton_3.clicked.connect(self.load_config)
        # 设置显示框编辑权限，禁止双击编辑数据值，选择
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView.setSelectionMode(QAbstractItemView.NoSelection)
        self.red_light()
        self.mail_host = ""
        self.mail_user = ""
        self.mail_pass = ""
        self.sender = ''
        self.receivers = []
        self.email_addresses = []
        self.email_string = ""
        self.output_data = []
        self.port = 25
        self.body_content = """ """
        self.subject = """ """
        # 配置文件
        # 初始化配置文件对象
        self.config = configparser.ConfigParser()

    def start_on_push_button_clicked(self):
        self.red_light()
        self.progressBar.setValue(0)
        # 获取lineEdit的文本值
        self.mail_host = self.lineEdit.text()
        self.mail_user = self.lineEdit_2.text()
        self.mail_pass = self.lineEdit_3.text()
        self.sender = self.lineEdit_4.text()
        self.subject = self.lineEdit_5.text()
        self.port = int(self.lineEdit_6.text())
        self.body_content = self.textEdit.toPlainText()
        self.email_string = self.textEdit_2.toPlainText()

        if(self.mail_host == '' or self.mail_user == '' or self.mail_pass == '' or self.sender == '' or self.subject == '' or self.body_content == '' or self.email_string == '' or self.port == None):
            select = QMessageBox.warning(self, "错误", "存在未填写的信息，或信息为空", QMessageBox.Yes,
                                         QMessageBox.Yes)
        else:
            # 创建多线程，并将value传递给time_consuming_task方法
            self.green_light()
            worker = Start(self.start_thread)
            # 在后台线程中执行耗时任务
            QThreadPool.globalInstance().start(worker)
    #开始线程
    def start_thread(self):
        # 以分号为分隔符拆分邮箱地址
        self.email_addresses = self.email_string.split(';')
        num_addresses = len(self.email_addresses)
        self.progressBar.setValue(0)  # 设置进度条的最小值
        self.progressBar.setMaximum(int(num_addresses))  # 设置进度条的最大值
        i = 1
        # 逐个发送邮件
        for email in self.email_addresses:
            message = MIMEText(self.body_content, 'plain', 'utf-8')
            message['From'] = self.mail_user
            message['To'] = str(email)
            message['Subject'] = Header(self.subject, 'utf-8')
            smtpObj = smtplib.SMTP(self.mail_host, self.port)
            smtpObj.connect(self.mail_host, self.port)
            # set_debuglevel(1)时SMTP 对象会打印与邮件服务器的交互过程，包括与服务器的命令和响应，以及交换的数据内容，用于debug
            smtpObj.set_debuglevel(0)
            smtpObj.starttls()
            smtpObj.login(self.mail_user, self.mail_pass)
            smtpObj.sendmail(self.sender, email, message.as_string())
            print(email+"邮件发送成功")
            current_time = datetime.now().strftime('%H:%M:%S')
            self.add_string_to_listView("[I]"+ email + "邮件发送成功，时间：" + current_time)
            smtpObj.quit()
            self.progressBar.setValue(i)
            i += 1
        print("所有邮件发送完成")
        self.red_light()

    # 绿灯
    def green_light(self):
        self.label_led1.setStyleSheet(
            '''QLabel{background-color: rgb(0, 255, 0);border-radius: 25px; border: 3px groove gray;border-style: outset;}''')

    # 红灯
    def red_light(self):
        self.label_led1.setStyleSheet(
            '''QLabel{background-color: rgb(255, 0, 0);border-radius: 25px; border: 3px groove gray;border-style: outset;}''')

    def set_listView_content(self):
        # 使用QStringListModel来设置listView的数据
        model = QStringListModel()
        model.setStringList(self.output_data)
        self.listView.setModel(model)

    def add_string_to_listView(self, new_string):
        self.output_data.append(new_string)
        # 更新 listView 的内容
        self.set_listView_content()

    # 保存变量到配置文件
    def save_config(self):
        file_path, _ = QFileDialog.getSaveFileName(None, '保存配置文件', '',
                                                   'ini files (*.ini)')
        config = configparser.RawConfigParser()
        self.mail_host = self.lineEdit.text()
        self.mail_user = self.lineEdit_2.text()
        self.mail_pass = self.lineEdit_3.text()
        self.sender = self.lineEdit_4.text()
        self.subject = self.lineEdit_5.text()
        self.body_content = self.textEdit.toPlainText()
        self.email_string = self.textEdit_2.toPlainText()
        self.port = int(self.lineEdit_6.text())
        # 将各个变量的值以 JSON 格式保存到配置文件中
        config['DEFAULT'] = {
            'mail_host': str(self.mail_host),
            'mail_user': str(self.mail_user),
            'mail_pass': str(self.mail_pass),
            'sender': str(self.sender),
            'subject': str(self.subject),
            'body_content': self.body_content,
            'email_string': self.email_string,
            'port': str(self.port),
        }
        if(file_path == ''):
            self.add_string_to_listView("[E]发送账号" + self.mail_user + "配置保存失败")
        else:
            with open(file_path, 'w') as configfile:
                config.write(configfile)
                self.add_string_to_listView("[I]发送账号" + self.mail_user + "配置保存完成")

    # 加载配置文件中的值
    def load_config(self):
        file_path, _ = QFileDialog.getOpenFileName(None, '打开配置文件', '',
                                                   'ini files (*.ini)')
        if (os.path.exists(file_path)):
            config = configparser.ConfigParser()
            config.read(file_path)
            # 将配置文件中的值加载到对应的变量中
            if 'DEFAULT' in config:
                self.mail_host = str(config['DEFAULT'].get('mail_host', ''))
                self.mail_user = str(config['DEFAULT'].get('mail_user', ''))
                self.mail_pass = str(config['DEFAULT'].get('mail_pass', ''))
                self.sender = str(config['DEFAULT'].get('sender', ''))
                self.subject = str(config['DEFAULT'].get('subject', ''))
                self.body_content = str(config['DEFAULT'].get('body_content', ''))
                self.email_string = str(config['DEFAULT'].get('email_string', ''))
                self.port = int(config['DEFAULT'].get('port', ''))
            self.lineEdit.setText(self.mail_host)
            self.lineEdit_2.setText(self.mail_user)
            self.lineEdit_3.setText(self.mail_pass)
            self.lineEdit_4.setText(self.sender)
            self.lineEdit_5.setText(self.subject)
            self.lineEdit_6.setText(str(self.port))
            self.textEdit.setText(self.body_content)
            self.textEdit_2.setText(self.email_string)
            self.add_string_to_listView("[I]发送账号" + self.mail_user + "配置加载完成")
        else:
            self.add_string_to_listView("[E]发送账号" + self.mail_user + "配置加载失败")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    MainWindow1 = LoginClass()
    MainWindow1.show()
    sys.exit(app.exec_())