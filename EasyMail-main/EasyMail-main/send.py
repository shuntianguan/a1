import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
subject = "sB!"
msg = "你是智障吗!"
from_address = "2947708947@qq.com"
to_address = "3751185634@qq.com"
password = "jdsjlfkyizgfdgha"  # 使用授权码，而不是邮箱密码

# 创建邮件
message = MIMEText(msg, 'plain')
message['From'] = from_address
message['To'] = to_address
message['Subject'] = subject
attachment = open("C:\Users\86135\Desktop\电子2024计网汇报.xlsx", 'rb')  
part = MIMEBase('application', 'octet-stream')
part.set_payload(attachment.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', f'attachment; filename={file_path.split("/")[-1]}')
# 发送邮件
try:
    with smtplib.SMTP_SSL('smtp.qq.com', 465) as server:
        server.login(from_address, password)
        server.sendmail(from_address, to_address, message.as_string())
    print("邮件发送成功")
except Exception as e:
    print("邮件发送成功")
