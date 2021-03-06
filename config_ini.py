# -------------------------------------------------------------------
# Модуль для работы с ini-файлом
# 25.01.18
# -------------------------------------------------------------------
import configparser, os

path = 'json_f.ini'

def create_config(path):
	"""
	Создать конфиг с установками по-умолчанию
	"""
	config = configparser.ConfigParser()
	config.add_section("Settings")
	config.set("Settings", "file", "docs.json")
	config.set("Settings", "age", "7")
	config.set("Settings", "log", "logfile.txt")
	config.set("Settings", "login", "kost@polimer.vn.ua")
	config.set("Settings", "password", "simple_pass")
	config.set("Settings", "from", "info@polimer.vn.ua")
	config.set("Settings", "host", "10.0.1.50")
	config.set("Settings", "port", "25")
	config.set("Settings", "pause", "20")
	config.set("Templates", "message", "message.html")
	config.set("Templates", "message_ret", "message_ret.html")
	config.set("Templates", "message_empty", "message_empty.html")
		
	with open(path, "w") as config_file:
		config.write(config_file)
 
def get_config(path):
	"""
	Прочитать конфиг
	"""
	if not os.path.exists(path):
		create_config(path)
	
	config = configparser.ConfigParser()
	config.read(path)
	return config
	
def get_all_settings(path):
	"""
	Прочитать весь ini подряд
	"""
	config = get_config(path)
	a = {}
	for each_section in config.sections():
		for each_key, each_val in config.items(each_section):
			a.update({each_key: each_val})
			#return each_key
			#return each_val
	for key in a.keys():
		print(key)

def get_setting(path, section, setting):
	"""
	Прочитать установки из конфига
	"""
	config = get_config(path)
	value = config.get(section, setting)
#	# Для отладки
#	msg = "{section} {setting} is {value}".format(section=section, setting=setting, value=value) 
#	print(msg)
#	# -----
	return value

def update_setting(path, section, setting, value):
	"""
	Обовить установки
	"""
	config = get_config(path)
	config.set(section, setting, value)
	with open(path, "w") as config_file:
		config.write(config_file)

def delete_setting(path, section, setting):
	"""
	Удалить установки
	"""
	config = get_config(path)
	config.remove_option(section, setting)
	with open(path, "w") as config_file:
		config.write(config_file)	

# Примеры вызова функций
#host = get_setting(path, 'Settings', 'host')
#age = get_setting(path, 'Settings', 'age')
#get_all_settings(path)
#print(get_all_settings(path))
#update_setting(path, "Settings", "age", "12")
#delete_setting(path, "Settings", "pause")
