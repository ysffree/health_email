# -*-coding: utf-8 -*-
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
import email
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
		#att = MIMEText(open(content['attach']).read(),'utf-8')
		#att["Content-Disposition"] = 'attachment; filename="{0}"'.format(content["attach_name"])
		att = MIMEApplication(open(content['attach'], 'rb').read())
		#att = email.message_from_file(content['attach'])
		#att["Content-Type"] = 'application/octet-stream'
		att.add_header("Content-Disposition", "attachment", filename=('utf-8', '', content['attach_name']))
		#encoders.encode_base64(att)
		message.attach(att)
		self.conn.sendmail(self.sender, content['to'].split(';')+content['bcc'].split(';'), message.as_string())
	
	def __del__(self):
		self.conn.quit()
		self.conn.close()

if __name__=='__main__':
	content = [{'to': 'zhaowei@smartquerier.com', 'bcc': '990346855@qq.com', 'subject': '报告交付：SH180000053-张暖悉--智力四项基因检测', 'html': 'test', 'attach': '/fastzone/sfyang/dev/health_email/report/29-Jan-2019-16-57-32/SH180000053-张暖悉--智力四项基因检测.pdf', 'attach_name': 'SH180000053-张暖悉--智力四项基因检测.pdf'},
{'to': 'gqbu@smartquerier.com', 'bcc': 'zhaowei@smartquerier.com', 'subject': '报告交付：SH180003282-张涵博--智力四项基因检测', 'html':'test', 'attach': '/fastzone/sfyang/dev/health_email/report/29-Jan-2019-16-57-32/SH180003282-张涵博--智力四项基因检测.pdf', 'attach_name': 'SH180003282-张涵博--智力四项基因检测.pdf'},
{'to': '18352516825@163.com', 'bcc': 'zhaowei@smartquerier.com', 'subject': '报告交付：SH180002884-金皓言--智力四项基因检测', 'html':'test', 'attach': '/fastzone/sfyang/dev/health_email/report/29-Jan-2019-16-57-32/SH180002884-金皓言--智力四项基因检测.pdf', 'attach_name': 'SH180002884-金皓言--智力四项基因检测.pdf'}]
	sendemail = SendEmail('smtp.smartquerier.com','smarthealth@smartquerier.com','!Huisuan19')
	for con in content:
		sendemail.send_email(con)
	del sendemail
	





