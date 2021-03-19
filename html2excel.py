# -*- coding: utf-8 -*-
"""
Created on Sat Jan 30 01:13:01 2021

@author: minam
"""


import openpyxl
from openpyxl.styles import Font
from bs4 import BeautifulSoup
import PySimpleGUI as sg






def strip_colorcode(stylecode):
    a=stylecode.replace('color:','').replace(';','').replace('#','')
    a ="ff"+a

    return a


def html2excel(htmlpath,name):
    
    wb = openpyxl.Workbook()
    sheet = wb.active
    namepath=name+".xlsx"
    sheet.title=namepath
    
    #ログのファイルを開く(ファイル名や位置を変えたい場合は第一引数を変更)
    with open(htmlpath,mode='rt', encoding='utf-8') as f:
        #beautifulsoupのオブジェクトを生成
        soup = BeautifulSoup(f.read(),'html.parser')
    
        for comment in soup.find_all('p'):
            #1つ目spanを抽出、中の値を抜き出し代入(タブ) 
            tab=comment.find_all('span')[0].get_text(strip=True)
            #2つ目spanを抽出、中の値を抜き出し代入（名前）
            name=comment.find_all('span')[1].get_text(strip=True)
            #3つ目spanを抽出、中の値を抜き出し代入（内容）
            text=comment.find_all('span')[2].get_text(strip=True)
            #pタグのstyleを抜き出し、カラーコードのみに成形
            colorcode=strip_colorcode(comment['style'])
            #print(colorcode)
            
            acomment = [tab, name, text]
            sheet.append(acomment)
            current = sheet[sheet.max_row]
            for a in current:
                a.font=Font(color=colorcode)
            
        #ログのファイルを閉じる
        f.close()
        
    wb.save(namepath)


layout = [
    [sg.Text("html ファイル名"), sg.InputText(), sg.FileBrowse(key="file1", file_types=(("html　ファイル","*.html"),))],
    [sg.Text("保存ファイル名"), sg.InputText(key="new_file_name")],
    [sg.Submit(), sg.Cancel()],
]
window = sg.Window("html converter", layout)
event, values = window.read()


html2excel(values["file1"],values["new_file_name"])

window.close()
