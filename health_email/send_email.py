#-*-coding: utf-8 -*-
# @Author:Sifen Yang
# @Date: 2018-11-06 13:56:34
# @Last  Modified by: Sifen Yang
# @Last Modified time: 2018-11-10 08:40:32


# *************************************************************************
# email.py
#
# 发邮件的工具
# *************************************************************************


import smtplib
from email.header import Header
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class SendEmail(object):


	def __init__(self, host, username, password):
		self.conn = smtplib.SMTP_SSL(host)
		self.conn.login(username, password)
		#self.conn.set_debuglevel(1)
		self.sender = username

	def send_email(self, content):
		#创建一个带附件的实例
		message = MIMEMultipart()
		message['From'] = self.sender
		message['To'] = content['to']
		#message['Bcc'] = content['bcc']
		message['Subject'] = Header(content['subject'], 'utf-8')
		# 邮件正文内容
		message.attach(MIMEText(content['html'], 'html', 'utf-8'))
		# 构造附件
		att = MIMEApplication(open(content['attach'], 'rb').read())
		att.add_header("Content-Disposition", "attachment", filename=('gbk', '', content['attach_name']))
		encoders.encode_base64(att)
		message.attach(att)
		self.conn.sendmail(self.sender, content['to'].split(';')+content['bcc'].split(';'), message.as_string())
	
	def __del__(self):
		self.conn.quit()
		self.conn.close()







