# -------------------------------------------------------------------
# Программа для рассылки сообщений контрагенам о долгах по документам 
# ООО-предприятия "Полимер"
# 01/06/18 ver. 0.1 by Kostya Pakhomov
# Python3 required
# -------------------------------------------------------------------
import json, smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_ADDRESS = 'kost@polimer.vn.ua'
SPECIAL_ADDRESS = 'info@polimer.vn.ua'
PASSWORD = 'kostas46'
MY_HOST = '10.0.1.50'
MY_PORT = 25

def get_contacts(filename):
	
	number_pp = []		# Порядковый номер записи
	client_names = []	# Список для названий организаций
	docum = []			# Список документов (секция)
	docum_n = []		# Список номеров документа
	docum_date = []		# Спсиок дат документов
	docum_desc = []		# Список описаний документов
	docum_invoice = []	# Список счетов на оплату
	managers = []		# Список для имен менеджеров (затем email'ов)
	manager_emails = []	# Список для e-mail'ов менеджеров
#	manager_names = []	# Список для имен менеджеров (для копии предыдущего)
	contact_infos = []	# Список для секции контрагентов
	contact_emails = []	# Список для e-mail'ов контрагентов

	for item in (json.load(open(filename))):
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
	
#	manager_names = list(managers) # Делаем копию списка имен менеджеров в новый список

	# Добываем данные из раздела "Документ"
	for item in docum:
		try:
			docum_n.append(item['Номер документа'])
			docum_date.append(item['Дата документа'])
			docum_desc.append(item['Описание'])
			docum_invoice.append(item['Счет на оплату'])
		except KeyError: # Если нет данных - заносим N/A
			docum_n.append(item['N/A'])
			docum_date.append(item['N/A'])
			docum_desc.append(item['N/A'])
			docum_invoice.append(item['N/A'])

	# Добываем данные из раздела "Контактная информация"
	for item in contact_infos:
		try:
			contact_emails.append(item['Адрес электронной почты контрагента для обмена электронными документами'])
		except KeyError: # Если нет данных - заносим N/A
			contact_emails.append('N/A')

#	# Поменяем в списке managers e-mail'ы на ФИО менеджеров
#	for n,i in enumerate(managers):
#		if i == 'Лінник Н.О.':
#			managers[n] = 'natali@polimer.vn.ua'
#		if i == 'Рябчинська Л.Г.':
#			managers[n] = 'lilyam@polimer.vn.ua'
#		if i == 'Пивоваров В.Д.':
#			managers[n] = 'slava@polimer.vn.ua'
#		if i == 'Марчук О.О.':
#			managers[n] = 'moleg@polimer.vn.ua'
#		if i == 'Плахотнюк Ю. І.':
#			managers[n] = 'julia@polimer.vn.ua'
#		if i == 'Захарчук В.С.':
#			managers[n] = 'valera@polimer.vn.ua'
		
#	with open(file_write, 'w') as f:		
#		for i in range(len(client_names)):
#			with open(file_write, 'a') as f:
#				f.write(client_names[i])
#				f.write(" ")
#				f.write(contact_emails[i])
#				f.write(" ")
#				f.write(manager_emails[i])
#				f.write("\n")
	return number_pp, client_names, docum_n, docum_date, docum_desc, docum_invoice, contact_emails, managers, manager_emails

def read_template(filename):
	with open(filename, 'r', encoding='utf-8') as template_file:
		template_file_content = template_file.read()
	return Template(template_file_content)

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
	nums, names, docum_n, docum_date, docum_desc, docum_invoice, contact_emails, managers, manager_emails = get_contacts('path copy.json')
	message_template = read_template('message.txt')
	message_ret_template = read_template('message_ret.txt')
	message_empty_template = read_template('message_empty.txt')

	for nums, name, doc_n, doc_date, doc_desc, doc_invoice, con_email, man_name, man_email in zip(nums, names, docum_n, docum_date, docum_desc, docum_invoice, contact_emails, managers, manager_emails):

		if man_email == "":
			message_empty = message_empty_template.substitute(NUMBER_PP=nums, CLIENT_NAME=name)
			print(message_empty)
			send_emails(MY_ADDRESS, SPECIAL_ADDRESS, message_empty)
		elif con_email == "N/A":
			message_ret = message_ret_template.substitute(NUMBER_PP=nums, CLIENT_NAME=name, MANAGER_NAME=man_name)
			print(message_ret)
			send_emails(MY_ADDRESS, man_email, message_ret)
		else:
			message = message_template.substitute(NUMBER_PP=nums, CLIENT_NAME=name, DOC_NUMBER=doc_n, DOC_DATE=doc_date, DOC_DESC=doc_desc.title(), DOC_INVOICE=doc_invoice, MANAGER_NAME=man_name, MANAGER_EMAIL=man_email)
			print(message)
			send_emails(man_email, con_email, message)

if __name__ == '__main__':
	main()