import  urllib.request, json, urllib.parse, requests, urllib3, json, cchardet, random, datetime, re, tkinter, traceback
from requests.packages.urllib3.util.retry import Retry
from urllib3.exceptions import InsecureRequestWarning
from requests.adapters import HTTPAdapter
from lxml import html
from urllib.parse import urlparse
from tkinter import *
import tkinter as tk
import tkinter.ttk as ttk
from urllib3.exceptions import InsecureRequestWarning
urllib3.disable_warnings(InsecureRequestWarning)

def convert_to_url(base, target_path):
    component = urlparse(base)
    directory = re.sub('/[^/]*$', '/',component.path)
    url_return = ""
    while(True):
        #[0] 絶対パスのケース（簡易版)
        petrn1 = re.compile("^http").search(target_path)
        petrn2 = re.compile("^\/\/.+").search(target_path)
        petrn3 = re.compile("^\/[^\/].+").search(target_path)
        petrn4 = re.compile("^\.\/(.+)").search(target_path)
        petrn5 = re.compile("^([^\.\/]+)(.*)").search(target_path)        
        petrn6 = re.compile("^([^\.\/]+)(.*)").search(target_path)    
        if petrn1: url_return = target_path; break
        if petrn2: url_return = component.scheme + ":" + target_path; break # [1]「//exmaple.jp/aa.jpg」のようなケース
        if petrn3: url_return = component.scheme + "://" + component.hostname + target_path; break # [2]「/aaa/aa.jpg」のようなケース
        if petrn4: url_return =  component.scheme + "://" + component.hostname + directory  + re.findall("^\.\/(.+)", target_path)[0]; break # [3]「./aa.jpg」のようなケース
        if petrn5:
            match_find = re.findall("^([^\.\/]+)(.*)", target_path)
            url_return =  component.scheme + "://" + component.hostname + directory + match_find[0][0] + match_find[0][1]; break
        # [5]「../aa.jpg」のようなケース
        if petrn6:
            match_find = re.findall("\.\./", target_path)
            nest = len(match_find)
            dir_name = re.sub('/[^/]*$', '/',component.path) + "\n"
            dir_array = dir_name.split('/')
            dir_array.pop(0)
            dir_array.pop()
            dir_count = len(dir_array)
            count = dir_count - nest
            pathto=""
            i = 0
            while i < count:
                pathto += "/" + dir_array[i];
                i += 1 
            file = target_path.replace("../","")    
            url_return =  component.scheme + "://" + component.hostname + pathto + "/" + file
            break
    return url_return

def get_item_url(dic, txt_shousai, making_type):
    url = dic['URL']
    item_urls = []
    log_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    txt_shousai.insert(tkinter.END, f"[{log_time}]：ページネーション\n")
    txt_shousai.see("end")
    count = 0
    while True:
        try:
            count += 1
            txt_shousai.insert(tkinter.END, f"{url}\n")
            txt_shousai.see("end")
            item_url, page_url = [], []
            base_url = dic['URL']
            pagenation_xpath =  dic['ページネーション']
            item_xpath = dic["アイテム要素"]
            get_html = lambda lxml, item_xpath: html.fromstring(lxml.content).xpath(item_xpath)
            retries = Retry(total = 3, backoff_factor = 1, status_forcelist = [ 500, 502, 503, 504, 443 ])
            requests.Session().mount('https://', HTTPAdapter(max_retries = retries))
            requests.Session().mount('http://', HTTPAdapter(max_retries = retries))
            responce = requests.get(url, verify = False, timeout=10)     
            responce.encoding = cchardet.detect(responce.content)["encoding"]
            page_lxml = get_html(responce, pagenation_xpath)
            if page_lxml:
                page_url = [convert_to_url(base_url, lxml.attrib['href']) for lxml in page_lxml][0]

            item_lxml = get_html(responce, item_xpath)
            if item_lxml:
                item_url = [convert_to_url(base_url, lxml.attrib['href']) for lxml in item_lxml]

            url = page_url
            item_urls.append(item_url)
            if not url:
                break
                
            if making_type == "テスト" and count > 3:
                break
            
        except:
            log_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            txt_shousai.insert(tkinter.END, f"[{log_time}]：ページネーションエラー\n【原因】リクエストエラーで途中で終了\n【確認】\n")
            break
            
    item_urls = [item_url  for item in item_urls for item_url in item]
    return item_urls

def get_responce_item(txt_shousai, item_urls):
    item_contents = []
    count = 0
    for url in item_urls: 
        try:
            count += 1
            txt_shousai.insert(tkinter.END, f"{url}\n")
            txt_shousai.see("end")
            get_html = lambda lxml: html.fromstring(lxml.content)
            retries = Retry(total = 3, backoff_factor = 1, status_forcelist = [ 500, 502, 503, 504, 443 ])
            requests.Session().mount('https://', HTTPAdapter(max_retries = retries))
            requests.Session().mount('http://', HTTPAdapter(max_retries = retries))
            responce = requests.get(url, verify = False, timeout=10)        
            responce.encoding = cchardet.detect(responce.content)["encoding"]
            lxml = get_html(responce)
            item_contents.append({"リクエスト":lxml,"URL":url})
        except:
            pass
        
    return item_contents


def main(item_content, dic):
    get_text =lambda soup, xpath: " ".join([item.text_content() for item in soup.xpath(xpath) if item.text_content() is not None]) if str(type(xpath)) ==  "<class 'str'>" else None
    soup = item_content["リクエスト"]
    url = item_content["URL"]
    column_1 = get_text(soup,dic["カラム①"]) if dic["カラム①"] else None
    column_2 = get_text(soup,dic["カラム②"]) if dic["カラム②"] else None
    column_3 = get_text(soup,dic["カラム③"]) if dic["カラム③"] else None
    column_4 = get_text(soup,dic["カラム④"]) if dic["カラム④"] else None
    column_5 = get_text(soup,dic["カラム⑤"]) if dic["カラム⑤"] else None
    column_6 = get_text(soup,dic["カラム⑥"]) if dic["カラム⑥"] else None
    column_7 = get_text(soup,dic["カラム⑦"]) if dic["カラム⑦"] else None
    column_8 = get_text(soup,dic["カラム⑧"]) if dic["カラム⑧"] else None
    column_9 = get_text(soup,dic["カラム⑨"]) if dic["カラム⑨"] else None
    column_10 = get_text(soup,dic["カラム⑩"]) if dic["カラム⑩"] else None
    tage_lst = [url, column_1, column_2,column_3,column_4,column_5,column_6,column_7, column_8,column_9,column_10]
    return tage_lst
