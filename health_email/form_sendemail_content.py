#! /usr/bin/env python3.6
# -*- coding: utf-8 -*-
# @Author: Saifen Yang
# @Date: 2019-01-10
# @Last Modified by: Saifen Yang
# @Last Modified time: 2018-01-10


import os
import pandas as pd
import jinja2


def email_content(dir_path, template_name):
	# 从excel中取出对应样本的地址
	for file in os.listdir(dir_path):
		if os.path.splitext(file)[1] in ['.xls','.xlsx']:
			detect_project, sampleid_address = get_address(os.path.join(dir_path, file))
	content_generator = get_content(detect_project, sampleid_address, dir_path, template_name)
	return content_generator

def get_address(excel_file):
	data = pd.read_excel(excel_file, sheet_name=0, keep_default_na=False)
	detect_project = data['检测项目'][0].strip()
	sampleid_address = {}
	if detect_project == '智力四项基因检测':
		for i in range(len(data)):
			sampleid = data['样本编号'][i].strip()
			address = {'to':data['邮箱'][i].strip(),'bcc':data['暗送邮箱（对应讲师邮箱）'][i].strip()}
			sampleid_address[sampleid] = address
	if detect_project == '套餐B： 女性体质健康综合评估':
		detect_project = detect_project.split('：')[1]
		for i in range(len(data)):
			sampleid = data['样本编号'][i].strip()
			address = {'to':data['电子邮箱'][i].strip(),'bcc':''}
			sampleid_address[sampleid] = address
	return detect_project, sampleid_address

def get_content(detect_project, sampleid_address, dir_path, template_name):
	for file in os.listdir(dir_path):
		if os.path.splitext(file)[1] == '.pdf':
			attach = os.path.join(dir_path, file)
			file_name = os.path.splitext(file)[0]
			sampleid = file_name.split('-')[0]
			subject = '报告交付：' + file_name
			customer_name = file_name.split('-')[1]
			to = sampleid_address[sampleid]['to']
			bcc = sampleid_address[sampleid]['bcc']
			html  = render_template(template_name, customer_name, detect_project)
			content = {
			'to': to,
			'bcc': bcc,
			'subject': subject,
			'html': html,
			'attach': attach,
			'attach_name': file
			}
			yield content


def render_template(template_name, customer_name, detect_project):
	env = jinja2.Environment(
		loader = jinja2.PackageLoader('health_email', 'templates')
	)
	template = env.get_template(template_name)
	return template.render(customer_name=customer_name, detect_project=detect_project)





