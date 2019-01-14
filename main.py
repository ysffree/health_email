# -*- coding: utf-8 -*-
# @Author: Saifen Yang
# @Date: 2019-01-10
# @Last Modified by: Saifen Yang
# @Last Modified time: 2018-01-10


import os
import pprint
import traceback
from configparser import ConfigParser
from health_email.download_email_content import DownloadEmail
from health_email.form_sendemail_content import email_content
from health_email.send_email import SendEmail





def main():
	try:
		cfg = ConfigParser()
		config_file = os.path.join(os.path.dirname(__file__),'config','email.ini')
		cfg.read(config_file)
		imap_host = cfg.get('email', 'imap_host')
		smtp_host = cfg.get('email', 'smtp_host')
		username = cfg.get('email', 'username')
		password = cfg.get('email', 'password')
		backup = cfg.get('path', 'backup')
		template_name = cfg.get('path', 'template_name')
		download_email = DownloadEmail(imap_host, username, password)
		dir_list = download_email.download_attach_from_mail(base_dir=backup)
		del download_email
		if not dir_list:
			return
		send_email = SendEmail(smtp_host, username, password)
		for data_dir in dir_list:
			content_generator = email_content(data_dir, template_name)
			for content in content_generator:
				#pprint.pprint(content)
				send_email.send_email(content)
		del send_email
	except:
		message = '\n' + traceback.format_exc()
		print (message)
		
			
if __name__ == '__main__':
	main()		



