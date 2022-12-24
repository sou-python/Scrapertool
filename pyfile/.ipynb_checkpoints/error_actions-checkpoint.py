import datetime, re

def main(num_start, dicts):
    num = num_start -1
    log_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")   
    eroor = []
    for dic in dicts:
        num += 1
        columns_need = list(dic.keys())[:4]
        for columns in columns_need:
            if str(type(dic[columns])) !=  "<class 'str'>":
                eroor.append(f"[{log_time}]：【空白エラー】\n{num}行目：{columns}空白しないで")

        columns_xpath = list(dic.keys())[7:]
        for columns in columns_xpath:
            if str(type(dic[columns])) == "<class 'str'>":
                if re.findall("(/[A-Z])",dic[columns]):
                    eroor.append(f"[{log_time}]：【大文字エラー】\n{num}行目：{dic[columns]}\n'/'のあとに大文字使用しないで")
    return eroor