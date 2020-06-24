# -------------------------------------------------------------------
# Программа для рассылки сообщений контрагенам о долгах по документам
# из JSON-файла, генерируемого 1С Предприятием
# (включает рассылку PDF-вложений)
# ООО-предприятие "Полимер"
# 03.09.18
# -------------------------------------------------------------------
import configparser, datetime, os, os.path, json, re, smtplib, time
import config_ini as ini
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

__author__ = "Kostya Pakhomov"
__version__ = "0.5.2"
__email__ = "naluster@gmail.com"

# Сегодняшняя дата и время
today_date = datetime.datetime.today()
today_date_printable = datetime.datetime.now().strftime("%d.%m.%Y in %H:%M:%S")

# Читаем json_f.ini
JSON_FILE = ini.get_setting(ini.path, "Settings", "file")
BASE_AGE = int(ini.get_setting(ini.path, "Settings", "age"))
LOG_FILE = ini.get_setting(ini.path, "Settings", "log")
MY_ADDRESS = ini.get_setting(ini.path, "Settings", "login")
PASSWORD = ini.get_setting(ini.path, "Settings", "password")
SPECIAL_ADDRESS = ini.get_setting(ini.path, "Settings", "from")
PDF_FILES = ini.get_setting(ini.path, "Settings", "pdf_path")
MESSAGE_FILE = ini.get_setting(ini.path, "Templates", "message")
MESSAGE_FILE_RET = ini.get_setting(ini.path, "Templates", "message_ret")
MESSAGE_FILE_EMPTY = ini.get_setting(ini.path, "Templates", "message_empty")
MY_HOST = ini.get_setting(ini.path, "Settings", "host")
MY_PORT = ini.get_setting(ini.path, "Settings", "port")
MY_PAUSE = int(ini.get_setting(ini.path, "Settings", "pause"))

def is_file_exist(filename): # Проверка существования файла и его доступности для чтения
	if os.path.isfile(filename) and os.access(filename, os.R_OK):
		return filename

def get_contacts(filename):
	
	number_pp = []		# Порядковый номер записи
	client_name = []	# Список для названий организаций
	docum = []			# Список документов (секция)
	docum_n = []		# Список номеров документа
	docum_date = []		# Список дат документов
	docum_desc = []		# Список описаний документов
	docum_pdf = []		# Список файлов-накладных PDF
	docum_invoice = []	# Список счетов на оплату 
	manager = []		# Список для имен менеджеров (затем email'ов)
	manager_email = []	# Список для e-mail'ов менеджеров
	contact_infos = []	# Список для секции контрагентов
	contact_email = []	# Список для e-mail'ов контрагентов

	for item in (json.load(open(filename, encoding="utf8"))):
		try:
			number_pp.append(item['Номер п/п'])
			client_name.append(item['Наименование клиента'])
			manager.append(item['Основной менеджер'])
			if item['Эл. адрес менеджера'] == '':
				manager_email.append('alla@polimer.vn.ua')
			else:
				manager_email.append(item['Эл. адрес менеджера'])
			docum.append(item['Документ'])
			contact_infos.append(item['Контактная информация'])
		except KeyError: # Если нет данных - заносим N/A
			number_pp.append(['N/A'])
			client_name.append('N/A')
			managers.append('N/A')
			manager_email.append('N/A') 
			docum.append('N/A')
			contact_infos.append('N/A')

	# Добываем данные из раздела "Документ"
	for item in docum:
		try:
			docum_n.append(item['Номер документа'])
			docum_date.append(item['Дата документа'])
			docum_desc.append(item['Описание'])
			docum_pdf.append(item['Имя файла'])
			docum_invoice.append(item['Счет на оплату'])
		except KeyError: # Если нет данных - заносим N/A
			docum_n.append(item['N/A'])
			docum_date.append(item['N/A'])
			docum_desc.append(item['N/A'])
			docum_pdf.append(item['N/A'])
			docum_invoice.append(item['N/A'])

	# Поменяем в списке docum_desc описания на более удобочитаемые используя RegEx
	for n,i in enumerate(docum_desc):
		if re.findall(r'нет д\w{1,}', i, re.I):
			docum_desc[n] = 'Відсутні документи'
		if re.findall(r'дор\s{1,}\d{1,}', i, re.I):
			docum_desc[n] = 'Відсутнє доручення'
	# Поменяем в счетах на оплату текст с русского на украинский
	for n,i in enumerate(docum_invoice):
		if 'Счет' in docum_invoice[n]:
			a = docum_invoice[n].split()
			b = ' '.join(a[6:])
			docum_invoice[n] = 'Рахунок на оплату покупцю ' + a[4] + ' від ' + b
		
	# Добываем данные из раздела "Контактная информация"
	for item in contact_infos:
		try:
			contact_email.append(item['Адрес электронной почты контрагента для обмена электронными документами'])
		except KeyError: # Если нет данных - заносим N/A
			contact_email.append('N/A')

	return number_pp, client_name, docum_n, docum_date, docum_desc, docum_pdf, docum_invoice, contact_email, manager, manager_email

def read_template(filename):
	with open(filename, 'r', encoding='utf-8') as template_file:
		template_file_content = template_file.read()
	return Template(template_file_content)

def modification_date(filename):
	modified_date = datetime.datetime.fromtimestamp(os.path.getmtime(filename))
	duration = today_date - modified_date
	return duration.days > BASE_AGE

def send_emails(from_addr, to_addr, text, subj, pdf_file=None):

	# Создаем сообщение
	msg = MIMEMultipart('mixed') # Обязательно 'mixed', а не 'alternative', иначе не видно вложения pdf!
	# Установим параметры сообщения
	msg['From']=from_addr
	msg['To']=to_addr
	msg['Subject']=subj
	msg.add_header('Content-Type','text/html')
	# Добавим тело сообщения (plain or html)
	msg.attach(MIMEText(text, 'html')) # HTML
	
	# Вложение (Base64)
	if pdf_file != None: 
		# Если передан параметр 'pdf_file'
		# Отображаемое имя во вложении
		filename = pdf_file+'.pdf'
		# Открыть файл с диска
		attachment = open(PDF_FILES+'/'+pdf_file+'.pdf', 'rb')
		# Закодировать в Base64 и прикрепить к письму
		part = MIMEBase('application', 'octet-stream')
		part.set_payload((attachment).read())
		encoders.encode_base64(part)
		part.add_header('Content-Disposition', 'attachment; filename = %s' % filename)
		msg.attach(part)
	
	# Подключимся к SMTP серверу
	server = smtplib.SMTP(host=MY_HOST, port=MY_PORT)
	server.starttls()
	server.login(MY_ADDRESS, PASSWORD)
	# Отправим сообщение через сервер, установленный ранее
	server.sendmail(from_addr, to_addr, msg.as_string())
	# Завершим соединение с сервером и выйдем
	server.quit()

def main():
	
	count = 0 # Счетчик сообщений
	subject = ['Сервісне повідомлення', 'Отчет о работе json_f']
	
	if is_file_exist(JSON_FILE):
		if modification_date(JSON_FILE):
			with open(LOG_FILE, 'a') as f:
				f.write('%s - File is older than %s days!\n' % (today_date_printable, BASE_AGE))
			msg = '{today_date} - Файл с данными на {BASE_AGE} дней старше, чем нужно!'.format(today_date=today_date_printable, BASE_AGE=BASE_AGE)
			print(msg)
			send_emails(SPECIAL_ADDRESS, MY_ADDRESS, msg, subject[1])
		else:
			# Message Template
			message_file = open(MESSAGE_FILE, 'r')
			messages = message_file.read()
			# Empty Message Template
			message_empty_file = open(MESSAGE_FILE_EMPTY, 'r')
			messages_empty = message_empty_file.read()
			# Returning Message Template
			message_ret_file = open(MESSAGE_FILE_RET, 'r')
			messages_ret = message_ret_file.read()

			for number_pp, client_name, docum_n, docum_date, docum_desc, docum_pdf, docum_invoice, contact_email, manager, manager_email in zip(*get_contacts(JSON_FILE)):
				print(number_pp, client_name, docum_n, docum_date, docum_desc, docum_pdf, docum_invoice, contact_email, manager, manager_email)	
#				if manager_email == "":
#					message_empty = {'NUMBER_PP': number_pp, 'CLIENT_NAME': client_name}
#					print(messages_empty % message_empty)
#					send_emails(MY_ADDRESS, SPECIAL_ADDRESS, (messages_empty % message_empty), subject[0])
#					message_empty_file.close()
				if contact_email == "N/A":
					message_ret = {'NUMBER_PP': number_pp, 'CLIENT_NAME': client_name, 'MANAGER_NAME': manager}
					print(messages_ret % message_ret)
					send_emails(MY_ADDRESS, manager_email, (messages_ret % message_ret), subject[0])
					message_ret_file.close()
				else:
					message = {'NUMBER_PP': number_pp, 'CLIENT_NAME': client_name, 'DOC_NUMBER': docum_n, 'DOC_DATE': docum_date, 'DOC_DESC': docum_desc, 'DOC_INVOICE': docum_invoice, 'DOC_PDF': docum_pdf, 'MANAGER_NAME': manager, 'MANAGER_EMAIL': manager_email}
					print(messages % message)
					send_emails(manager_email, contact_email, (messages % message), docum_desc, docum_pdf)
					message_file.close()
				count = count + 1
				time.sleep(MY_PAUSE) # пауза
			with open(LOG_FILE, 'a') as f:
				f.write('%s - %s messages was sent successfully!\n' % (datetime.datetime.now().strftime("%d.%m.%Y in %H:%M:%S"), count))
			msg = 'Сегодня, {now_date} было успешно разослано {count} писем должникам документов.\n'.format(now_date=datetime.datetime.now().strftime("%d.%m.%Y in %H:%M:%S"), count=count)
			print(msg)
			send_emails(SPECIAL_ADDRESS, MY_ADDRESS, msg, subject[1])
	else:
		print("No JSON file was found!")
		with open(LOG_FILE, 'a') as f:
			f.write('%s - No JSON file was found! Check if it really exists!\n' % (today_date_printable))

if __name__ == '__main__':
	main()
