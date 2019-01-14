#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
# @Author: Saifen Yang
# @Date: 2019-01-10
# @Last Modified by: Saifen Yang
# @Last Modified time: 2018-01-10


#import一些库
import imaplib, email, os
# import pandas as pd
# import numpy as np
from datetime import datetime
from email.header import decode_header


class DownloadEmail(object):
    

    def __init__(self, host, username, password):
        self.conn = imaplib.IMAP4_SSL(host, '993')
        self.conn.login(username, password)

    def download_attach_from_mail(self, mailbox='INBOX', tag='UnSeen', base_dir=None):
        self.conn.select(mailbox)
        resp, mails = self.conn.search(None, tag)
        if mails == [b'']:
            return None
        dir_list = []
        for mail in mails:
            resp, data = self.conn.fetch(mail.split()[len(mail.split())-1],'(RFC822)')
            emailbody = data[0][1]
            content = email.message_from_bytes(emailbody)
            dir_name = self.__parse_head(content)
            dir_path = os.path.join(base_dir, dir_name)
            if not os.path.exists(dir_path):
                os.mkdir(dir_path)
            self.__parse_body(content, dir_path)
            dir_list.append(dir_path)
            self.__movemail(mail)
        return dir_list

    def __movemail(self, mail, new_mailbox='Finish'):
        self.conn.copy(mail, new_mailbox)
        self.conn.store(mail, '+FLAGS', '\\Deleted')

    def __parse_head(self, mail):
        date = email.header.decode_header(mail.get('date'))
        # 以收件时间为下载文件夹名文件夹名
        dir_name = date[0][0].split(',')[1].split('+')[0].replace(':','-').strip().replace(' ','-')
        return dir_name

    def __parse_body(self, mail, dir_path):
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            try:
                fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
            except:
                pass
            filePath = os.path.join(dir_path, fileName)
            if os.path.exists(filePath):
                os.remove(filePath)
            with open(filePath, 'wb') as f:
                f.write(part.get_payload(decode=True))

    def __del__(self):
        self.conn.close()
        self.conn.logout()

if __name__ == '__main__':
    aa = GetEmail('imap.mxhichina.com', 'sfyang@smartquerier.com', 'Ysf+839593')
    aa.download_attach_from_mail()
    del aa






# #连接到qq企业邮箱，其他邮箱调整括号里的参数
# conn = imaplib.IMAP4_SSL("imap.mxhichina.com",'993')

# #用户名、密码，登陆
# conn.login("sfyang@smartquerier.com","Ysf+839593")
# #选定一个邮件文件夹
# #可以用conn.list()查看都有哪些文件夹。中文的文件夹名称可能是乱码，没关系，直接拷贝过来就行了。
# conn.select("INBOX")

# #提取了文件夹中所有邮件的编号，search功能在本邮箱中没有实现……
# resp, mails = conn.search(None,'UnSeen')


# # #提取了指定编号（最新一封）的邮件
# resp, data = conn.fetch (mails[0],'(RFC822)')

# emailbody = data[0][1]

# mail = email.message_from_bytes(emailbody)

# date = email.header.decode_header(mail.get('date'))

# print()

# # fileName = '没有找到任何附件！' 
# #获取邮件附件名称
# for part in mail.walk():
#     if part.get_content_maintype() == 'multipart':
#         continue
#     if part.get('Content-Disposition') is None:
#         continue
#     fileName = part.get_filename()  


# #如果文件名为纯数字、字母时不需要解码，否则需要解码
#     try:
#         fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
#     except:
#         pass

# #如果获取到了文件，则将文件保存在制定的目录下
#     if fileName != '没有找到任何附件！':
#         filePath = os.path.join("D:\work\市场data\慧算医疗\healt_eamil", fileName)

#         if not os.path.isfile(filePath):
#             fp = open(filePath, 'wb')
#             fp.write(part.get_payload(decode=True))
#             fp.close()
#             print ("附件已经下载，文件名为：" + fileName)
#         else:
#             print ("附件已经存在，文件名为：" + fileName)
# conn.close()
# conn.logout()
