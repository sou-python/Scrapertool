import configparser

inifile = configparser.SafeConfigParser()
inifile.read('config/settings.ini', "UTF-8")

# FILE
EXCEL_PATH = inifile.get('EXCEL', 'PATH')
SAVE_EXCEL_PATH = inifile.get('EXCEL', 'SAVE_PATH')