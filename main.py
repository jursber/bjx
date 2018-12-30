# -*- coding: utf-8 -*-
'''
---------------------------------------
北极星网站爬虫
Author:构造线
Email:jursber@163.com
---------------------------------------
'''

from tkinter import *
from tkinter import ttk
import tkinter.font as tkFont
import requests
import urllib.parse
from bs4 import BeautifulSoup
import xlwt
import os
import threading
import time
import random

comb_switch = ('电力要闻', '售电政策', '风电政策', '光伏政策', '储能政策', '输配电政策', '软件政策')

switch_url = {
    '电力要闻': r'http://news.bjx.com.cn/zt.asp?topic=',
    '售电政策': r'http://fd.bjx.com.cn/NewsList?id=100&page=',
    '风电政策': r'http://fd.bjx.com.cn/NewsList?num=4&id=100&page=',
    '光伏政策': r'http://guangfu.bjx.com.cn/NewsList.aspx?typeid=100&page=',
    '输配电政策': r'http://shupeidian.bjx.com.cn/NewsList?id=100&page=',
    '储能政策': r'http://chuneng.bjx.com.cn/NewsList?id=100&page=',
    '软件政策': r'http://xinxihua.bjx.com.cn/list?catid=100&page=',
    '火电政策': r'http://huodian.bjx.com.cn/NewsList?num=2&id=100&page='
}

switch_headers = {
    '电力要闻': 'news.bjx.com.cn',
    '售电政策': 'shoudian.bjx.com.cn',
    '风电政策': 'fd.bjx.com.cn',
    '光伏政策': 'guangfu.bjx.com.cn',
    '输配电政策': 'shupeidian.bjx.com.cn',
    '储能政策': 'chuneng.bjx.com.cn',
    '软件政策': 'xinxihua.bjx.com.cn',
    '火电政策': 'huodian.bjx.com.cn'
}
headers = {
    'User-Agent':
    'Mozilla/5.0 (Windows NT 10.0; win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'\
    ' Chrome/64.0.3282.140 Safari/537.36 MicroMessenger/6.5.2.501 NetType/'\
    'WIFI WindowsWechat QBCore/3.43.691.400 QQBrowser/9.0.2524.400',
}
    
start = 1
end = 100
t=4
page_num=0

def crawler_news(path, target, key_words, start_page=1,end_page=100):  #新闻爬虫主函数
    start_time = time.time()
    data_list, sheet_list = [], {}
    global headers,switch_url,switch_headers
    
    headers['Referer'] = switch_headers[target]
    headers['Host'] = switch_headers[target]

    for word in key_words:
        key_word_code = urllib.parse.quote(word.encode('gb2312'))
        url_head = switch_url[target] + key_word_code + r'&page='
        out_print('正在解析 电力要闻-'+word+' 的页面数据...')
        end_judge=['first']
        for i in range(start_page, end_page + 1):
            url = url_head + str(i)
            try:
                req = requests.get(url, headers=headers)
            except:
                out_print('由于抓取过于频繁，你已被封禁，请稍后再试！！！！！')
                return
            if req.status_code != 404:
                out_print('    正在爬取 ' + word + ' 的第' + str(i) + '页，已耗时：' + str(round(time.time() - start_time, 2)) + 's')
                soup = BeautifulSoup(req.text, 'html.parser')
                
                title_list = [each.get_text() for each in soup.select('ul.list_left_ztul > li > a')]
                href_list = [each.get('href') for each in soup.select('ul.list_left_ztul > li > a')]
                add_data_list = [each.get_text() for each in soup.select('ul.list_left_ztul > li > span')]
                
                if (title_list==[] ) or (end_judge[0]==title_list[0] and end_judge[-1]==title_list[-1]):
                    global page_num
                    page_num=i
                    break
                                
                for title, href, add_data in zip(title_list, href_list,add_data_list):
                    data = {'title': title, 'href': href, 'add_data': add_data}
                    data_list.append(data)
                delay(t)  
                end_judge=title_list[:]
            else:
                break
        sheet_list[word] = data_list
        data_list = []
    out_print('----------------爬取完成！----------------')
    delay(1)
    heads = ['编号', '标题', '链接', '发布日期']
    workbook = xlwt.Workbook()
    out_print('开始写入数据到Excel：')
    delay(1)
    for keyword in sheet_list:
        out_print('    正在写入' + keyword + '到Excel...')
        delay(1)
        worksheet = workbook.add_sheet(keyword)
        i = 0
        for head in heads:
            worksheet.write(0, i, head)
            i += 1
        i = 1
        for data in sheet_list[keyword]:
            worksheet.write(i, 0, i)
            worksheet.write(i, 1, data['title'])
            worksheet.write(i, 2, data['href'])
            worksheet.write(i, 3, data['add_data'])
            i += 1
    try:
        workbook.save(path + '\\' + '电力要闻' + time.strftime('%Y-%m-%d', time.localtime()) + '.xls')
    except:
        out_print('文件写入错误，请检查权限！')
        return
    delay(1)
    end_time = time.time()  #结束时间
    out_print('----------------写入完成！----------------')

    out_print('总计用时：' + str(round(end_time - start_time, 2)) + 's')
    return

def crawler_policy(path, target, key_words, start_page=1,end_page=100):  #政策爬虫主函数
    start_time = time.time()  #记录开始时间
    data_list = []
    global headers,switch_headers,switch_url
    
    headers['Referer'] = switch_headers[target]
    headers['Host'] = switch_headers[target]
    out_print('正在解析 '+ target +' 页面数据...')
    end_judge=['first']
    
    for i in range(start_page, end_page + 1):
        #网页链接
        url = switch_url[target] + str(i)
        # 网页的headers
        try:
            req = requests.get(url, headers=headers)
        except:
            out_print('由于抓取过于频繁，你已被封禁，请稍后再试！！！！！')
            return
        code = req.status_code  # 404停止
        if code != 404:  # 获取内容
            out_print('    正在爬取第' + str(i) + '个网页，当前耗时：' + str(round(time.time() - start_time, 2)) + 's')
            soup = BeautifulSoup(req.text, 'html.parser')
            
            title_list = [each.get_text() for each in soup.select('ul.list_left_ul > li > a')]
            href_list = [each.get('href')for each in soup.select('ul.list_left_ul > li > a')]
            add_time_list = [each.get_text() for each in soup.select('ul.list_left_ul > li > span')]
            
            if (title_list==[] ) or (end_judge[0]==title_list[0] and end_judge[-1]==title_list[-1]):
                global page_num
                page_num=i
                break
                             
            for title, href, add_time in zip(title_list, href_list,
                                             add_time_list):
                data = {'title': title, 'href': href, 'add_time': add_time}
                data_list.append(data)  # 信息打包
            end_judge=title_list[:]
        else:
            break
        delay(t)
    book = xlwt.Workbook()
    sheet1 = book.add_sheet(target, cell_overwrite_ok=True)  # 写入excel
    out_print('----------------爬取完成！----------------')
    
    
    heads = ['编号', '标题', '链接', '发布日期']
    delay(1)
    out_print('开始写入'+ target +'到excel...')
    ii = 0
    for head in heads:
        sheet1.write(0, ii, head)
        ii += 1
    i = 1
    for data in data_list:
        sheet1.write(i, 0, i)
        sheet1.write(i, 1, data['title'])
        sheet1.write(i, 2, data['href'])
        sheet1.write(i, 3, data['add_time'])
        i += 1
    try:
        book.save(path + '\\' + target +time.strftime('%Y-%m-%d', time.localtime()) + '.xls')
    except:
        out_print('文件写入错误，请检查权限！')
        return
    end_time = time.time()  #结束时间
    out_print('----------------写入完成！----------------')
    out_print('总计用时：' + str(round(end_time - start_time, 2)) + 's')
    return

def comb_edit_able(*args):
    if comb1.get() == '电力要闻':
        ent_key_words['state'] = 'normal'
    else:
        ent_key_words['state'] = 'disabled'

def val_checking(*args):#检查输入是否有问题
    #back_door=output.get(END)
    #if back_door.find('end')!= -1:
    #    end=int(back_door[back_door.find('end')+3:])
    #else:
    output.delete(0.0, END)
    root.update()
    start_button.update()
    save_path = ent_path.get()
    target_comb = comb1.get()
    key_words = ent_key_words.get().split(',')
    
    symbol = 1
    if os.path.exists(save_path) != True:
        output.insert(END, '    输入的文件夹不存在！\n')
        output.update()
        symbol = 0
    if target_comb not in comb_switch:
        output.insert(END, '    请选择正确的类别！ \n')
        output.update()
        symbol = 0
    if target_comb == '电力要闻' :
        if key_words == [''] :
            out_print('   关键字不能为空，多个关键字用英文逗号隔开！ \n ')
            symbol = 0
        elif len(key_words)>5:
            out_print('   为避免被封，关键字最多只能添加5个！\n')
            symbol = 0
    if symbol == 1:
        crawler(save_path, target_comb, key_words, start, end)
    return
def crawler(path, target, key_words=None, start_page=1, end_page=100):#选择执行主函数
    ent_path['state'] = 'disabled'
    start_button['state'] = 'disabled'
    comb1['state'] = 'disabled'
    ent_key_words['state'] = 'disabled'
    if target == '电力要闻':
        crawler_news(path, target, key_words, start_page, end_page)
    else:
        crawler_policy(path, target, key_words, start_page, end_page)
    start_button['state'] = 'normal'
    comb1['state'] = 'normal'
    ent_key_words['state'] = 'normal'
    ent_path['state'] = 'normal'
def out_print(s):#输出到文本框
    output.insert(END, s+'\n')
    output.update()
    output.see(END)
def delay(t,t0=0):#延时，慢点爬
    time.sleep(random.uniform(t0,t))

def main_fun():#多线程运行
    th=threading.Thread(target=val_checking)
    th.setDaemon(True)#守护线程
    th.start()
#-------------------主窗口-----------------------------------------------------
root = Tk()
root.title("北极星电力网爬虫V1.0  by:构造线")
root.geometry('403x303+750+260')
root.resizable(0, 0)
ctl_padx = 2
ft = tkFont.Font(family='Microsoft YaHei', size=9)
Label(root, text='请输入文件存储地址：', font=ft).grid(row=0, sticky=W)

ent_path = Entry(root, width=56, font=ft)
ent_path.grid(row=1, padx=ctl_padx, columnspan=3, sticky=W)
ent_path.insert(END, 'C:\\users\\Administrator\\Desktop')

Label(root, text='请选择查询类别：', font=ft).grid(column=0, row=2, sticky=W)
Label(root, text='请输入电力要闻关键字：', font=ft).grid(column=1, row=2, sticky=W)

comb1 = ttk.Combobox(root, font=ft)
comb1.grid(row=3, column=0, sticky=W, padx=ctl_padx)
comb1['value'] = comb_switch

ent_key_words = Entry(root, font=ft, state='disabled')
ent_key_words.grid(row=3, column=1, sticky=W, padx=ctl_padx)

start_button = Button(
    root, text='开始爬取', font=ft, width=10, height=2, relief=GROOVE)
start_button.grid(row=2, rowspan=3, column=2, sticky=W, padx=4)

fm = Frame(width=58, height=205)
fm.grid(
    row=6, columnspan=3, sticky=W + E + N + S, padx=ctl_padx, pady=ctl_padx)
fm.grid_propagate(0)

output = Text(
    fm,
    font=ft,
    background='white',
    width=56,
    height=13,
    padx=ctl_padx,
    pady=ctl_padx)
output.grid(sticky=N + E + S + W)

tips='\n       --------------------北极星电力网爬虫V1.0--------------------\n\n'\
'                                为避免网站封禁当前IP\n\n'\
'                       请不要在短时间内频繁使用本工具！\n\n'\
'           如有任何疑问，请联系软件作者：'\
'jursber@163.com\n\n'\
'       ----------------------------------------------------------------\n'
output.insert(END,tips)
#窗口事件
#--------------------------------------------------------------------------
start_button['command'] = main_fun
comb1.bind("<<ComboboxSelected>>", comb_edit_able)

root.mainloop()