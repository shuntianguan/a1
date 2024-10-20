from SignInWindow import Ui_Dialog as SignInWindow_Ui
from MainWindow import Ui_MainWindow as MainWindow_Ui
from PyQt5 import QtWidgets
from PyQt5.QtCore import QThread, pyqtSignal, Qt
import sys
from mailserver import *
import traceback
import os
import base64
import poplib
from email import policy
from email.parser import BytesParser
from email.header import decode_header, make_header
import sql
import pymysql
from snownlp import SnowNLP
from deep_translator import GoogleTranslator

def translate_text(text, dest_language='zh-CN'):
    translated = GoogleTranslator(target=dest_language).translate(text)
    return translated

def analyze_chinese_content(content):
    s = SnowNLP(content)
    # 得到情感分数，范围从 0 到 1，接近 1 表示积极，接近 0 表示消极
    score = s.sentiments
    print(score)
    # 根据情感分析分数判断是否为钓鱼邮件
    return score < 0.1 or score > 0.9# 自定义阈值
zongzong=[]
sqlli=sql.SQL()
filejiji=[]
mail_ind=-1
mailil=' '
#重发送专用
resender=[]
# 定义EmailClient类，用于连接到POP3服务器并从指定的邮件地址获取邮件
class EmailClient:
    # 在初始化函数中，设置POP3服务器的来源、用户、密码和待查询的目标邮件地址
    def __init__(self, host, user, password):
        self.pop_server = poplib.POP3_SSL(host,995)  # 使用POP3协议通过SSL安全连接到邮件服务器
        self.pop_server.user(user)  # 输入用户邮箱
        self.pop_server.pass_(password)  # 输入用户邮箱密码
       # self.target_email = target_email  # 输入待查询的目标邮件地址

    # 定义一个函数，用以清除文件名中的无效字符
    def sanitize_folder_name(self, name):
        invalid_characters = "<>:\"/\\|?*@"
        for char in invalid_characters:  # 遍历所有无效字符
            name = name.replace(char, "_")  # 将无效字符替换为下划线
        return name  # 返回清理后的名称

    # 定义一个函数，用以提取邮件的payload（有效载荷，即邮件主体内容）
    def get_payload(self, email_message,zhixing=0):
        if email_message.is_multipart():  # 判断邮件是否为多部分邮件
            for part in email_message.iter_parts():  # 如果是，则遍历其中的每一部分
                content_type = part.get_content_type()  # 获取该部分的内容类型
                if content_type == 'text/html':  # 如果内容类型为HTML，则返回该部分内容
                    return part.get_content()
                elif content_type == 'text/plain':  # 如果内容类型为纯文本，则返回该部分内容
                    return part.get_content()
                elif content_type == 'application/pdf' and zhixing==1:  # 如果内容类型为PDF，则返回该部分内容

                    save_directory = '../EasyMail-1/test2'
                    os.makedirs(save_directory, exist_ok=True)  # 创建目录（如果不存在）

                    filename = part.get_filename() or 'default.pdf'
                    file_path = os.path.join(save_directory, filename)

                    # 然后再写入文件
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    #return part.get_content()

        elif email_message.get_content_type() == 'text/html':  # 如果邮件非多部分形式，且为HTML类型，则返回邮件内容
            return email_message.get_content()
        elif email_message.get_content_type() == 'text/plain':  # 如果邮件非多部分形式，且为纯文本类型，则返回邮件内容
            return email_message.get_content()
        elif email_message.get_content_type() == 'application/pdf':  # 如果内容类型为PDF，则返回该部分内容
            return email_message.get_content()
    def num_email(self):
        return len(self.pop_server.list()[1])
    # 定义一个函数，用以获取邮件信息
    def fetch_email(self,shu=-1):
        num_emails = len(self.pop_server.list()[1])  # 获取邮箱内的邮件数量
        recv_email=[]
        # 遍历每一封邮件
        for i in range(num_emails):
            # 获取邮件内容
            response, lines, octets = self.pop_server.retr(i + 1)  # retr函数返回指定邮件的全部文本
            email_content = b'\r\n'.join(lines)  # 将所有行连接成一个bytes对象
            
            # 解析邮件内容
            email_parser = BytesParser(policy=policy.default)  # 创建一个邮件解析器
            email = email_parser.parsebytes(email_content)  # 解析邮件内容，返回一个邮件对象
            #print(email)
            # 解析邮件头部信息并提取发件人信息
            email_from = email.get('From').strip()  # 获取发件人信息，并去除尾部的空格
            email_from = str(make_header(decode_header(email_from)))  # 解码发件人信息，并将其转换为字符串
            #if email_from == self.target_email:  # 如果发件人地址与指定的目标邮件地址一致，对邮件进行处理
                # 解析邮件时间
            email_time = email.get('Date')  # 获取邮件时间

                # 提取邮件正文
            if i==shu:
                email_body = self.get_payload(email,1)  # 获取邮件正文
            else :
                email_body = self.get_payload(email)
            #print(email_from,email_time,email_body)
                #return email_body, email_time  # 返回邮件正文和时间
            recv_email.append([email_from,email_time,email_body])
        return recv_email
        #print("No new emails from", self.target_email)  # 如果没有从目标邮件地址收到新邮件，打印相应信息
        #return None, None  # 返回None
#client = EmailClient('pop.qq.com', '2947708947@qq.com', 'jdsjlfkyizgfdgha', '同程旅行 <tcwlsp@tcmail.17usoft.com>')
#client.fetch_email()


class SignInWindowUi(SignInWindow_Ui, QtWidgets.QDialog):
    def __init__(self):
        super(SignInWindowUi, self).__init__()
        BussLogic.win_stage=0
        self.setupUi(self)
        self.mail_server_address = "smtp.qq.com" # Default
        self.comboBoxServerAddress.activated.connect(self.select_server_address)

    def fetch_info(self):
        username = self.lineEditUsername.text()
        password = self.lineEditPassword.text()
        return self.mail_server_address, username, password

    def select_server_address(self):
        if self.comboBoxServerAddress.currentText() == "QQ Mail":
            self.mail_server_address = "smtp.qq.com"
        elif self.comboBoxServerAddress.currentText() == "NEU E-Mail":
            self.mail_server_address = "smtp.neu.edu.cn"
        elif self.comboBoxServerAddress.currentText() == "Gmail":
            self.mail_server_address = "smtp.google.com"
        elif self.comboBoxServerAddress.currentText() == "163 Mail":
            self.mail_server_address = "smtp.163.com"

class MainWindowUi(MainWindow_Ui, QtWidgets.QMainWindow):
    def __init__(self):
        super(MainWindowUi, self).__init__()
        self.setupUi(self)
        self.list_widget = self.listWidget
        self.stacked_widget = self.stackedWidget
        self.change_stacked_widget()

    def change_stacked_widget(self):
        self.list_widget.currentRowChanged.connect(self.display_subpage)

    def display_subpage(self, i):
        self.stacked_widget.setCurrentIndex(i)

    def set_text_edit(self, mail_server, username, password):
        text = mail_server + username + password
        self.textEdit.setText(text)

class MailServer(Smtp, Pop3):
    def __init__(self):
        self.mail_server = " "
        self.username = " "
        self.password = " "
        self.smtp = None
        self.pop3 = None

    def info_init(self, mail_server, username, password):
        self.mail_server = mail_server
        self.username = username
        self.password = password
        try:
            self.smtp = Smtp(mailserver=mail_server, username=username, password=password)
            self.pop3 = Pop3(mailserver=mail_server, username=username, password=password)
        except Exception as e:
            print(e)
        else:
            print(mail_server, username, password)

class BussLogic:
    def __init__(self):
        self.sign_in_window = SignInWindowUi()
        self.main_window = MainWindowUi()
        self.mail_server = MailServer()
        self.sender_proc = None
        self.sign_in_window.pushButtonSignin.clicked.connect(self.click_sign_in)
        self.main_window.pushButtonSend.clicked.connect(self.click_send)
        self.main_window.pushButtonSave.clicked.connect(self.click_save)
        self.main_window.pushButtonfujian.clicked.connect(self.fujian)
        self.main_window.Reflesh_Button.clicked.connect(self.reflesh_recv)
        self.main_window.listWidget.itemClicked.connect(self.show_recv)
        self.main_window.listWidget_2.itemClicked.connect(self.show_mail)
        self.main_window.listWidget.itemClicked.connect(self.show_draft)
        self.main_window.listWidget_3.itemClicked.connect(self.open_draft)
        self.main_window.pushButtonBackToInbox.clicked.connect(self.compose_to_inbox)
        self.main_window.pushButtonComposeAnotherEmail.clicked.connect(self.back_to_comp)
        self.main_window.Delete_Button.clicked.connect(self.delete_mail)    #pop 协议接收到的mail无法修改，编程下载相关pdf文件
        self.main_window.Resend_button.clicked.connect(self.resend_butt)
        self.main_window.pushButtontrans.clicked.connect(self.trans)
    def resend_butt(self):
        print('111')
        '''
        uid=self.uid
        result=sql.SQL.search_sql_by_uid_with_sender(self.mail_server.username,uid,'Draft')
        mail=Mail()
        mail.sender=result[0][0]
        mail.receiver=result[0][1]
        mail.topic=result[0][2]
        mail.uid=result[0][3]
        self.mail_server.smtp.mail = mail
        self.mail_server.smtp.path = "C:\\MailServer\\Draft"
        no=self.mail_server.smtp.sendmail()
        if(no==errno):
            self.main_window.close()
            self.sign_in_window.exec()
        sql.SQL.delete_sql(uid,'Draft')
        os.remove("C:\\MailServer\\Draft\\"+uid+'.txt')
        '''
        print(resender)
        self.sender_proc = Sender_proc(sender=self.mail_server.username,
                                       receiver=resender[0],
                                       subject=resender[2],
                                       message=resender[1])
        print('it is here')
        mail = self.sender_proc.gene_mailclass()
        self.sender_proc.store()
        global filejiji
        self.mail_server.smtp.filejihe=filejiji
        self.mail_server.smtp.mail = mail
        self.mail_server.smtp.path = "C:\\MailServer\\Draft"
        no=self.mail_server.smtp.sendmail()
        if(no==errno):
            self.main_window.close()
            self.sign_in_window.exec()
        
        filejiji=[]
        self.main_window.display_subpage(3)
        
    def delete_mail(self):
        '''
        uid=self.uid
        result=sql.SQL.search_sql_by_uid(self.mail_server.username,uid,'Mail')
        index=result[0][4]
        self.mail_server.pop3.recvmail('DELE',index)
        self.reflesh_recv()
        '''
        self.main_window.listWidget_2.clear()
            #定义一个EmailClient对象
        eyeyey=EmailClient('pop'+self.mail_server.mail_server[4:],self.mail_server.username,self.mail_server.password)
        global mail_ind
        jieshou=eyeyey.fetch_email(mail_ind)
        
        mail_ind=-1
        zongzong=[]
        for i in range(eyeyey.num_email()):
            zongzong=[]
            tmp=List_item(jieshou[i][1],jieshou[i][0],jieshou[i][2],i+1)
            self.main_window.listWidget_2.addItem(tmp)
            zongzong.append(tmp)
            self.main_window.listWidget_2.setItemWidget(tmp,tmp.widgit)
    def show_draft(self,item):
        if(item.text()=='Drafts'):
            '''
            self.main_window.listWidget_3.clear()
            recvaddr=self.mail_server.username
            dbtuple=sql.SQL.show_tables()
            dblist=''
            for i in dbtuple:
                dblist+=i[0]
            if(dblist.find('draft')==-1):
                sql.SQL.create_sql('Draft')
            recvtuple=sql.SQL.search_sql_by_sender(recvaddr,'Draft')
            if(recvtuple==None):
                return
            for t in recvtuple:
                tmp=List_item(t[2],t[0],t[3],t[4])
                self.main_window.listWidget_3.addItem(tmp)
                self.main_window.listWidget_3.setItemWidget(tmp,tmp.widgit)
            '''
            
            self.main_window.listWidget_3.clear()
            dbtuple=sqlli.getshuju()
            print(dbtuple)
            for i in range(len(dbtuple)):
                tmp=List_item(dbtuple[i][1],dbtuple[i][0],dbtuple[i][2],i)
                self.main_window.listWidget_3.addItem(tmp)
                self.main_window.listWidget_3.setItemWidget(tmp,tmp.widgit)
    def open_draft(self,item):
        '''
        uid=item.uid
        self.uid=uid
        f=open('C:\\MailServer\\Draft\\'+uid+'.txt')
        msg=f.read()
        self.main_window.textBrowser_2.setText(msg)
        '''
        global resender
        resender=[]
        if len(resender)== 0:
            resender.append(item.sender)
            resender.append(item.body)
            resender.append(item.subject)
        else :
            resender[0]=item.sender
            resender[1]=item.body
            resender[2]=item.subject
        print(resender)
        self.main_window.textBrowser_2.setText(item.body)

    def compose_to_inbox(self):
        self.main_window.stacked_widget.setCurrentIndex(1)
        
        
    def back_to_comp(self):
        self.main_window.lineEditTo.clear()
        self.main_window.lineEditSubject.clear()
        self.main_window.textEdit.clear()
        self.main_window.stacked_widget.setCurrentIndex(0)

    def click_sign_in(self):
        self.mail_server = MailServer()
        mail_server_address, username, password = self.sign_in_window.fetch_info()
        if(username=="" or password==""):
            return
        self.mail_server.info_init(mail_server=mail_server_address, username=username, password=password)
        BussLogic.win_stage=1
        self.sign_in_window.close()

    def click_send(self):
        self.sender_proc = Sender_proc(sender=self.mail_server.username,
                                       receiver=self.main_window.lineEditTo.text(),
                                       subject=self.main_window.lineEditSubject.text(),
                                       message=self.main_window.textEdit.toPlainText())
        mail = self.sender_proc.gene_mailclass()
        self.sender_proc.store()
        self.mail_server.smtp.mail = mail
        self.mail_server.smtp.path = "C:\\MailServer\\Draft"
        global filejiji
        self.mail_server.smtp.filejihe=filejiji
        no=self.mail_server.smtp.sendmail()
        if(no==errno):
            self.main_window.close()
            self.sign_in_window.exec()
        
        filejiji=[]
        #清空列表
        self.main_window.listWidget_fu.clear()
        self.main_window.display_subpage(3)
        
    def click_save(self):
        self.sender_proc = Sender_proc(sender=self.mail_server.username,
                                       receiver=self.main_window.lineEditTo.text(),
                                       subject=self.main_window.lineEditSubject.text(),
                                       message=self.main_window.textEdit.toPlainText())
        mail=self.sender_proc.gene_mailclass()
        self.sender_proc.store()
        #sql.SQL.add_sql(mail.sender,mail.receiver,mail.topic,mail.uid,0,'Draft')
        try:
            conn = pymysql.connect(host='localhost',user='root',password='LZPbs697...',database='mydatabase233')
        except Exception as e:
            print(f'数据库连接失败：{e}')
        cursor = conn.cursor()
        
        sql = "INSERT INTO draft (sendfrom, sendtime, sendbody) VALUES (%s, %s, %s)"
        val = (self.main_window.lineEditTo.text(), self.main_window.lineEditSubject.text(), self.main_window.textEdit.toPlainText())
        cursor.execute(sql, val)
        
        conn.commit()
        print('it is here')
        print(cursor.rowcount, "record inserted.")
        cursor.close()
        conn.close()
    def fujian(self):
        window = QtWidgets.QWidget()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(window, "选择文件")
        if file_path:
            print(f"选择的文件: {file_path}")
        global filejiji
        filejiji.append(file_path)
        self.main_window.listWidget_fu.clear()
        tmp=List_item(file_path,'附件','附件3',1)
        self.main_window.listWidget_fu.addItem(tmp)
        self.main_window.listWidget_fu.setItemWidget(tmp,tmp.widgit)
    def trans_info(self):
        mail_server_address, username, password = self.sign_in_window.fetch_info()
        self.mail_server.info_init(mail_server_address, username, password)

    def reflesh_recv(self):
        '''
        self.main_window.listWidget_2.clear()
        self.mail_server.pop3.recvmail('LIST',0)
        recvaddr=self.mail_server.username
        recvtuple=sql.SQL.search_sql(recvaddr,'Mail')
        if(recvtuple==None):
            return
        for t in recvtuple:
            tmp=List_item(t[2],t[0],t[3],t[4])
            self.main_window.listWidget_2.addItem(tmp)
            self.main_window.listWidget_2.setItemWidget(tmp,tmp.widgit)
        '''
        self.main_window.listWidget_2.clear()
            #定义一个EmailClient对象
        eyeyey=EmailClient('pop'+self.mail_server.mail_server[4:],self.mail_server.username,self.mail_server.password)
        jieshou=eyeyey.fetch_email()
        global mail_ind
        mail_ind=-1
        zongzong=[]
        for i in range(eyeyey.num_email()):
            zongzong=[]
            tmp=List_item(jieshou[i][1],jieshou[i][0],jieshou[i][2],i+1)
            self.main_window.listWidget_2.addItem(tmp)
            zongzong.append(tmp)
            self.main_window.listWidget_2.setItemWidget(tmp,tmp.widgit)
    def show_recv(self,item):
        
        if(item.text()=='Inbox'):
            '''
            self.main_window.listWidget_2.clear()
            recvaddr=self.mail_server.username
            print('it is here')
            dbtuple=sql.SQL.show_tables()
            
            dblist=''
            for i in dbtuple:
                dblist+=i[0]
            if(dblist.find('mail')==-1):
                sql.SQL.create_sql('Mail')
            recvtuple=sql.SQL.search_sql(recvaddr,'Mail')
            if(recvtuple==None):
                return
            for t in recvtuple:
                tmp=List_item(t[2],t[0],t[3],t[4])
                self.main_window.listWidget_2.addItem(tmp)
                self.main_window.listWidget_2.setItemWidget(tmp,tmp.widgit)
            '''
            self.main_window.listWidget_2.clear()
            #定义一个EmailClient对象
            eyeyey=EmailClient('pop'+self.mail_server.mail_server[4:],self.mail_server.username,self.mail_server.password)
            jieshou=eyeyey.fetch_email()
            global mail_ind
            mail_ind=-1
            zongzong=[]
            for i in range(eyeyey.num_email()):
                zongzong=[]
                tmp=List_item(jieshou[i][1],jieshou[i][0],jieshou[i][2],i+1)
                self.main_window.listWidget_2.addItem(tmp)
                zongzong.append(tmp)
                self.main_window.listWidget_2.setItemWidget(tmp,tmp.widgit)
    def show_mail(self,item):
        '''
        uid=item.uid
        self.uid=uid
        if(not os.path.exists('C:\\MailServer\\'+uid+'.txt')):
            self.mail_server.pop3.recvmail('RETR',item.index)
        f=open('C:\\MailServer\\'+uid+'.txt')
        msg=f.read()
        try:
            self.main_window.textBrowser.setText(msg)
        except Exception as e:
            print(e)
        '''
        global mail_ind
        mail_ind=item.index
        global mailil
        mailil=item.body
        
        self.main_window.textBrowser.setText(item.body)
        print(mailil)
        a=False
        if mailil!=None:
            a = analyze_chinese_content(item.body)
        
        if a :
            self.main_window.labelsense.setText("警告可能为钓鱼邮件")
        else:
            self.main_window.labelsense.setText("暂未检验出钓鱼信息")

    def trans(self):
        aaaa=translate_text(mailil)
        self.main_window.textBrowser.setText(aaaa)
class List_item(QtWidgets.QListWidgetItem):
    def __init__(self,subject,sender,uid,index):
        super().__init__()
        self.uid=uid
        self.index=index
        self.widgit=QtWidgets.QWidget()
        self.subject_label=QtWidgets.QLabel()
        self.subject_label.setText(subject)
        self.sender_label=QtWidgets.QLabel()
        self.sender_label.setText(sender)
        self.hbox=QtWidgets.QVBoxLayout()
        self.hbox.addWidget(self.sender_label)
        self.hbox.addWidget(self.subject_label)
        self.widgit.setLayout(self.hbox)
        self.setSizeHint(self.widgit.sizeHint())
        self.body=uid
        self.sender=sender
        self.subject=subject
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    buss = BussLogic()
    buss.sign_in_window.exec()
    # if(BussLogic.win_stage==0):
    #     exit()
    buss.main_window.show()
    sys.exit(app.exec_())
