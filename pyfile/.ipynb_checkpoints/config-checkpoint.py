import configparser

inifile = configparser.SafeConfigParser()
inifile.read('config/settings.ini', "UTF-8")

# SHEET
ITSZAI_SHEET_NAME = inifile.get('SHEET', 'ITSZAI_SHEET_NAME')
CMS_EC_SHEET_NAME = inifile.get('SHEET', 'CMS_EC_SHEET_NAME')
XPATH_SHEET_NANE = inifile.get('SHEET', 'XPATH_SHEET_NANE')
SPREADSHEET_KEY = inifile.get('SHEET', 'SPREADSHEET_KEY')
# FILE
KEY_FILE1 = inifile.get('FILE', 'KEY_FILE1')
KEY_FILE2 = inifile.get('FILE', 'KEY_FILE2')
PROXY_FILE = inifile.get('FILE', 'PROXY_FILE')
SUB_SUB_CATES_FILE = inifile.get('FILE', 'SUB_SUB_CATES_FILE')
