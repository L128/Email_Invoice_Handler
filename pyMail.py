# -*- coding: utf-8 -*-
import imaplib
import email
import sys
import os
import smtplib
import mimetypes
from email.mime import multipart, base, text
from email.encoders import encode_base64
from email.header import Header
import email.utils
import email.parser
import tempfile
import re

pattern_uid = re.compile(r'\d+ \(UID (?P<uid>\d+)\)')


def guessCharset(msg):
    charset = msg.get_charset()  # 获取编码方式
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()  # 获取内容类型字符串
        pos = content_type.find('charset=')  # 内容类型中查找“charset=”字符串的位置
        if pos >= 0:
            charset = content_type[pos + 8:].strip()  # 若存在上述字符串，则返回内容类型
    return charset


def decodeDict(msg, key):
    value, charset = email.header.decode_header(msg[key])[0]
    if charset:
        # decode if needed. Use guess if contradicts since charset can be 'unknown'.
        value = value.decode(guessCharset(msg) or charset)
    return value


def decodeStr(str):
    value, charset = email.header.decode_header(str)[0]
    if charset:
        value = value.decode(charset)
    return value


# *********接受邮件部分（IMAP）**********
# 处理接受邮件的类
class ReceiveMailDealer:

    # 构造函数(用户名，密码，imap服务器)
    def __init__(self, server, username, password):
        self.mail = imaplib.IMAP4_SSL(server)
        self.mail.login(username, password)
        self.select("INBOX")

    # 返回所有文件夹
    def showFolders(self):
        return self.mail.list()

    # 选择收件箱（如“INBOX”，如果不知道可以调用showFolders）
    def select(self, selector):
        return self.mail.select(selector)

    # 搜索邮件(参照RFC文档http://tools.ietf.org/html/rfc3501#page-49)
    def search(self, charset, *criteria):
        try:
            return self.mail.search(charset, *criteria)
        except Exception as e:
            print(e)
            self.select("INBOX")
            return self.mail.search(charset, *criteria)

    # 返回所有未读的邮件列表（返回的是包含邮件序号的列表）
    def getUnread(self):
        return self.search(None, "Unseen")

    # 返回所有收件箱列表（返回的是包含邮件序号的列表）
    def getInbox(self):
        return self.search(None, "Inbox")

    # 以RFC822协议格式返回邮件详情的email对象
    def getEmailFormat(self, num):
        data = self.mail.fetch(num, 'RFC822')
        if data[0] == 'OK' and data[1][0][0].decode().split()[1][1:] == 'RFC822':
            return email.message_from_bytes(data[1][0][1])
        else:
            return "fetch error"

    def getEmailUID(self, num):
        data = self.mail.fetch(num, "(UID)")
        return pattern_uid.match(data[1][0].decode()).group('uid')

    # 返回发送者的信息——元组（邮件称呼，邮件地址）
    @staticmethod
    def getSenderInfo(msg):
        return decodeDict(msg, "From")

    # 返回接受者的信息——元组（邮件称呼，邮件地址）
    @staticmethod
    def getReceiverInfo(msg):
        return decodeDict(msg, "To")

    # 返回邮件的主题（参数msg是email对象，可调用getEmailFormat获得）
    @staticmethod
    def getSubjectContent(msg):
        return decodeDict(msg, "subject")

    '''判断是否有附件，并解析（解析email对象的part）返回列表（内容类型，大小，文件名，数据流）'''
    @staticmethod
    def parse_attachment(message_part):
        content_disposition = message_part.get("Content-Disposition", '')
        if content_disposition:
            dispositions = content_disposition.strip().split(";")
            if bool(content_disposition and dispositions[0].lower() == "attachment"):
                file_data = message_part.get_payload(decode=True)
                attachment = {"content_type": message_part.get_content_type(), "size": len(file_data)}
                name = decodeStr(message_part.get_filename())
                attachment["name"] = name
                attachment["data"] = file_data
                # attachment["email_title"] = message_part.get
                '''保存附件
                fileobject = open(name, "wb")
                fileobject.write(file_data)
                fileobject.close()
                '''
                return attachment
        return None

    '''返回邮件的解析后信息部分
    返回列表包含（主题，纯文本正文部分，html的正文部分，发件人元组，收件人元组，附件列表）
    '''

    def getMailInfo(self, num):
        msg = self.getEmailFormat(num)
        attachments = []
        body = None
        # html = None
        html = []
        for part in msg.walk():
            attachment = self.parse_attachment(part)
            # print(part)
            if attachment:
                attachments.append(attachment)
            # elif part.get_content_type() == "text/plain":
            #     if body is None:
            #         body = ""
                # body += part.get_payload(decode=True)
            elif part.get_content_type() == "text/html":
                # if html is None:
                #     html = ""
                for links in str(part.get_payload(decode=True)).split():
                    if links.find('target="_blank">http') != -1:
                        html.append(links[links.index('>')+1: links.index('<')])
                # html += part.get_payload(decode=True)
        return {
            'subject': self.getSubjectContent(msg),
            # 'body': body,
            'html': html,
            'from': self.getSenderInfo(msg),
            'to': self.getReceiverInfo(msg),
            'attachments': attachments,
            'uid': self.getEmailUID(num)
            # 'label': label
        }


# *********发送邮件部分(smtp)************************************************************

class SendMailDealer:

    # 构造函数（用户名，密码，smtp服务器）
    def __init__(self, smtp, user, passwd, port=465, usettls=False):
        self.mailUser = user
        self.mailPassword = passwd
        self.smtpServer = smtp
        self.smtpPort = port
        self.mailServer = smtplib.SMTP_SSL(self.smtpServer, self.smtpPort)
        self.mailServer.ehlo()
        if usettls:
            self.mailServer.starttls()
        self.mailServer.ehlo()
        self.mailServer.login(self.mailUser, self.mailPassword)
        self.msg = multipart.MIMEMultipart()

    # 对象销毁时，关闭mailserver
    def __del__(self):
        self.mailServer.close()

    # 重新初始化邮件信息部分
    def reinitMailInfo(self):
        self.msg = multipart.MIMEMultipart()

    # 设置邮件的基本信息（收件人，主题，正文，正文类型html或者plain，可变参数附件路径列表）
    def setMailInfo(self, receiveUser, subject, texts, text_type, attachmentFilePaths):
        self.msg['From'] = self.mailUser
        self.msg['To'] = receiveUser

        self.msg['Subject'] = subject
        self.msg.attach(text.MIMEText(texts, text_type))
        for attachmentFilePath in attachmentFilePaths:
            self.msg.attach(self.getAttachmentFromFile(attachmentFilePath))

            # 自定义邮件正文信息（正文内容，正文格式html或者plain）

    def addTextPart(self, texts, text_type):
        self.msg.attach(text.MIMEText(texts, text_type))

    # 增加附件（以流形式添加，可以添加网络获取等流格式）参数（文件名，文件流）
    def addAttachment(self, filename, filedata):
        part = base.MIMEBase('application', "octet-stream")
        part.set_payload(filedata)
        encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % str(Header(filename, 'utf8')))
        self.msg.attach(part)

    # 通用方法添加邮件信息（MIMETEXT，MIMEIMAGE,MIMEBASE...）
    def addPart(self, part):
        self.msg.attach(part)

    # 发送邮件
    def sendMail(self):
        if not self.msg['To']:
            print("没有收件人,请先设置邮件基本信息")
            return
        self.mailServer.sendmail(self.mailUser, self.msg['To'], self.msg.as_string())
        print('Sent email to %s' % self.msg['To'])

        # 通过路径添加附件

    @staticmethod
    def getAttachmentFromFile(attachmentFilePath):
        part = base.MIMEBase('application', "octet-stream")
        part.set_payload(open(attachmentFilePath, "rb").read())
        encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % str(Header(attachmentFilePath, 'utf8')))
        return part
