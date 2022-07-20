emailSender = 'csywwang@comp.hkbu.edu.hk'                                     # 发件人邮箱账号
emailSenderPassword ='**********'                                 # 发件人邮箱密码
emailSenderName = "Yiwen"                                    # 发件人昵称
emailSMTPAddress = "mh2.comp.hkbu.edu.hk"                                   # 发件人邮箱SMTP地址（一般为smtp.邮箱后缀，如smtp.126.com）
emailSMTPPort = 465                                          # 发件人邮箱SMTP端口（非加密端口一般为25，加密端口一般为465）

emailTitle = "Yiwen的测试标题"                                      # 邮件主题（标题）
#emailContentFilename = "EmailContent.txt"                   # 邮件内容（文本形式）
emailContentFilename = "EmailContent.html"                 # 邮件内容（网页形式）

emailReceiversListFilename = "EmailReceiversList.csv"       # 收件人邮箱账号列表csv文件
failListFilename = "FailList.csv"                           # 发送失败的邮箱列表


import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart
from email.header import Header

# 读取收件人邮箱列表
emailReceiversList = open(emailReceiversListFilename, 'r', encoding="utf8").readlines()
emailReceivers = []
for each in emailReceiversList:    
    tmp = each.strip().split(",")[0]
    if tmp!="":
        emailReceivers.append(tmp)

# 读取邮件内容
if "html" in emailContentFilename:
    emailContentType = "html" 
else:
    emailContentType = "plain" 
emailContent = open(emailContentFilename, 'r', encoding="utf8").read()

# 连接服务器并发送邮件
failListFile = open(failListFilename, 'w', encoding="utf8")
try:
    server = smtplib.SMTP_SSL(emailSMTPAddress, emailSMTPPort)              # 发件人邮箱中的SMTP服务器
    server.login(emailSender, emailSenderPassword)                          # 发件人邮箱账号、邮箱密码
    
    successCount = 0
    for each in emailReceivers:                                             # 逐个邮箱发送，达到群发单显的效果
        try:
            message = MIMEMultipart()
            message['From'] = formataddr([emailSenderName, emailSender])
            message['Subject'] = emailTitle
            message['To'] =  each
            #邮件正文内容
            message.attach(MIMEText(emailContent, emailContentType, 'utf-8'))
            # 构造附件1，传送当前目录下的 test.txt 文件
            att1 = MIMEText(open('Example.png', 'rb').read(), 'base64', 'utf-8')
            att1["Content-Type"] = 'application/octet-stream'
            # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
            att1["Content-Disposition"] = 'attachment; filename="Example.png"'
            message.attach(att1)                                             # 对应收件人邮箱账号
            server.sendmail(emailSender, [each], message.as_string())           # 发件人邮箱账号、收件人邮箱账号、发送邮件
            print("成功发送邮件至："+each)
            successCount += 1
        except Exception:
            print("尝试发送至"+each+"失败")
            failListFile.write(each+"\n")
        
    server.quit()                                                           # 关闭与邮箱服务器的连接
    print("共有"+str(successCount)+"封邮件发送成功，"+str(len(emailReceivers)-successCount)+"封邮件发送失败")
except Exception: 
    print("与邮箱服务器连接失败")
failListFile.close()