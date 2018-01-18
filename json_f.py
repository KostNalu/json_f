# -------------------------------------------------------------------
# Программа для рассылки сообщений контрагенам о долгах по документам 
# ООО-предприятия "Полимер"
# 18.01.18 ver. 0.2 by Kostya Pakhomov
# Python3 required
# -------------------------------------------------------------------
import datetime, os, json, re, smtplib, time
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

BASE_AGE = 7							# JSON файл должен быть не старше чем ... дней 
JSON_FILE = 'docs.json'					# JSON файл
LOG_FILE = 'logfile.txt'				# LOG файл
MY_ADDRESS = 'kost@polimer.vn.ua'		# Адрес для логина на SMTP сервер
SPECIAL_ADDRESS = 'info@polimer.vn.ua'	# Спецадрес для отправки 
PASSWORD = ''					# Пароль для логина на SMTP сервер
MY_HOST = '10.0.1.50'					# SMTP сервер
MY_PORT = 25							# SMTP порт

def get_contacts(filename):
	
	number_pp = []		# Порядковый номер записи
	client_names = []	# Список для названий организаций
	docum = []			# Список документов (секция)
	docum_n = []		# Список номеров документа
	docum_date = []		# Список дат документов
	docum_desc = []		# Список описаний документов
#	docum_invoice = []	# Список счетов на оплату
	managers = []		# Список для имен менеджеров (затем email'ов)
	manager_emails = []	# Список для e-mail'ов менеджеров
	contact_infos = []	# Список для секции контрагентов
	contact_emails = []	# Список для e-mail'ов контрагентов

	for item in (json.load(open(filename, encoding="utf8"))):
		try:
			number_pp.append(item['Номер п/п'])
			client_names.append(item['Наименование клиента'])
			managers.append(item['Основной менеджер'])
			manager_emails.append(item['Эл. адрес менеджера'])
			docum.append(item['Документ'])
			contact_infos.append(item['Контактная информация'])
		except KeyError: # Если нет данных - заносим N/A
			number_pp.append(['N/A'])
			clent_names.append('N/A')
			managers.append('N/A')
			manager_emails.append('N/A')
			docum.append('N/A')
			contact_infos.append('N/A')
	
	# Добываем данные из раздела "Документ"
	for item in docum:
		try:
			docum_n.append(item['Номер документа'])
			docum_date.append(item['Дата документа'])
			docum_desc.append(item['Описание'])
#			docum_invoice.append(item['Счет на оплату'])
		except KeyError: # Если нет данных - заносим N/A
			docum_n.append(item['N/A'])
			docum_date.append(item['N/A'])
			docum_desc.append(item['N/A'])
#			docum_invoice.append(item['N/A'])

	# Поменяем в списке docum_desc описания на более удобочитаемые используя RegEx
	for n,i in enumerate(docum_desc):
		if re.findall(r'нет д\w{1,}', i):
			docum_desc[n] = 'Немає документів'
		if re.findall(r'дор\s{1,}\d{1,}', i):
			docum_desc[n] = 'Немає доручення'

	# Добываем данные из раздела "Контактная информация"
	for item in contact_infos:
		try:
			contact_emails.append(item['Адрес электронной почты контрагента для обмена электронными документами'])
		except KeyError: # Если нет данных - заносим N/A
			contact_emails.append('N/A')

	return number_pp, client_names, docum_n, docum_date, docum_desc, contact_emails, managers, manager_emails

def read_template(filename):
	with open(filename, 'r', encoding='utf-8') as template_file:
		template_file_content = template_file.read()
	return Template(template_file_content)

def modification_date(filename):
	today_date = datetime.datetime.today()
	modified_date = datetime.datetime.fromtimestamp(os.path.getmtime(filename))
	duration = today_date - modified_date
	return duration.days > BASE_AGE 

def send_emails(from_addr, to_addr, message):
	
	# Подключимся к SMTP серверу
	s = smtplib.SMTP(host=MY_HOST, port=MY_PORT)
	s.starttls()
	s.login(MY_ADDRESS, PASSWORD)

	msg = MIMEMultipart()       # Создаем сообщение

	# Установим параметры сообщения
	msg['From']=from_addr
	msg['To']=to_addr
	msg['Subject']='Повернення документів'
		
	# Добавим тело сообщения
	msg.attach(MIMEText(message, 'plain'))
		
	# Отправим сообщение через сервер, установленный ранее
	s.send_message(msg)
	del msg
		
	# Завершим соединение с сервером и выйдем
	s.quit()

def main():
	today_date = datetime.datetime.today()
	if modification_date(JSON_FILE):
		with open(LOG_FILE, 'a') as f:		
			f.write('%s: File is older than %s days!\n' % (today_date, BASE_AGE))
		print('%s: File is older than %s days!' % (today_date, BASE_AGE))
	else:
		nums, names, docum_n, docum_date, docum_desc, contact_emails, managers, manager_emails = get_contacts(JSON_FILE)
		message_template = read_template('message.txt')
		message_ret_template = read_template('message_ret.txt')
		message_empty_template = read_template('message_empty.txt')

		for nums, name, doc_n, doc_date, doc_desc, con_email, man_name, man_email in zip(nums, names, docum_n, docum_date, docum_desc, contact_emails, managers, manager_emails):

			if man_email == "":
				message_empty = message_empty_template.substitute(NUMBER_PP=nums, CLIENT_NAME=name)
				print(message_empty)
#				send_emails(MY_ADDRESS, SPECIAL_ADDRESS, message_empty)
			elif con_email == "N/A":
				message_ret = message_ret_template.substitute(NUMBER_PP=nums, CLIENT_NAME=name, MANAGER_NAME=man_name)
				print(message_ret)
#				send_emails(MY_ADDRESS, man_email, message_ret)
			else:
				message = message_template.substitute(NUMBER_PP=nums, CLIENT_NAME=name, DOC_NUMBER=doc_n, DOC_DATE=doc_date, DOC_DESC=doc_desc.title(), MANAGER_NAME=man_name, MANAGER_EMAIL=man_email)
				print(message)
#				send_emails(man_email, con_email, message)
			time.sleep(20) # пауза	

if __name__ == '__main__':
	main()
