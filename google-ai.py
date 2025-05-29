#!/usr/bin/python
# -*- coding: UTF-8 -*-
import requests
import pandas as pd
from numpy.core.defchararray import strip


"""
@:param path 文件路径 string
@:param sheet_name 表名 string
@:param column 要翻译的列名 如 A B C string 
@:param start 开始行 int
@:param end 结束行 int
@:param language 目标语言 string
"""
def process(
  path,
  sheet_name,
  column,
  start,
  end,
  language
):
  real_start = start - 2 if start >= 2 else 0
  real_end = end - 2 if end >= 2 else 0
  df = pd.read_excel(path, sheet_name=sheet_name)
  d_row = df.loc[real_start:real_end,column].values
  print(d_row)
  s_list = []
  translate_res= []
  for index in range(d_row.size):
    s = d_row[index]
    # print('当前: %d %s' % (index, d_column[index]))
    s_list.append(s)
    if len(s_list) == 100:
      res = do_translate('|'.join(s_list), language)
      print('结果: ', res)
      translate_res.extend(res)
      s_list.clear()

  if len(s_list) > 0:
    left_res = do_translate('|'.join(s_list), language)
    # left_res = ['Mehr', 'Datenschutzrichtlinie', 'Nutzungsbedingungen', 'Über uns', 'Versionsnummer: v$s', 'Bitte geben Sie das Sperrfeuer ein', 'Spezialeffekte', 'Farbe', 'Allgemein (Mehrfachauswahl)', 'blinkend \n']
    print("结果: ", left_res)
    translate_res.extend(left_res)
  print('translate res ', translate_res)
  write_res(df, translate_res, language, real_start, real_start + len(translate_res))
  df.to_excel(path, sheet_name=sheet_name, index=False, header=True)


def write_res(
  data_frame,
  res,
  language,
  start,
  end
):
  print('write res start: %d end: %d res size: %d' % (start, end, len(res)))
  try:
    column = data_frame[language]
  except:
    print('增加列: ', language)
    data_frame[language] = None
  for i in range(len(res)):
    print('set cell %d:%s %s' % (i, language, res[i]))
    data_frame.at[start + i, language] = strip(res[i])


"""
@:param text string
@:param language string
@:return list [string]
"""
def do_translate(text, language):
  api_key = "your api key"
  prompt = "you are a professional translator, I Have several words or sentence separated by |, translate '%s' into %s, keep the |, and only output the result, the content may contains some string placeholder, such in android development's %%1$s, or iOS or flutter,just keep the placeholder. If a word is a product name such as Google Play/Youtube/GMail, just keep it as English"
  headers = {'Content-Type': 'application/json'}
  proxies = {
    'http': 'http://127.0.0.1:7890',
    'https': 'http://127.0.0.1:7890'  # https -> http
  }
  s = (prompt % (text, language))
  # print('翻译: %s' % s)
  payloads = {
    'contents': [
      {
        'parts': [
          {
            'text': s
          }
        ]
      }
    ],
    "safetySettings": [
      {
        "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        "threshold": "BLOCK_NONE"
      },
      {
        "category": "HARM_CATEGORY_HARASSMENT",
        "threshold": "BLOCK_NONE"
      },
      {
        "category": "HARM_CATEGORY_HATE_SPEECH",
        "threshold": "BLOCK_NONE"
      },
      {
        "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
        "threshold": "BLOCK_NONE"
      },
    ],
  }
  r = requests.post(
    "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=%s" % api_key,
    json=payloads, headers=headers, proxies=proxies)
  if r.status_code == 200:
    data = r.json()
    candidates = data.get('candidates', [])
    if len(candidates) > 0:
      candidate = candidates[0]
      parts = candidate.get('content', {}).get('parts', [])
      if len(parts) > 0:
        text = parts[0].get('text', '')
        return text.split('|')
  else:
    print('translate Failed code: ', r.status_code)
  print(r.content)

print('start...')
file_path = input("请输入表格路径：")
sheet_name= input("请输入工作表名字如Sheet1：")
base_column = input("请输入要翻译列第一行的名字如English：")
start_end = input("请输入开始行和结束行如2:10：")
range_split = start_end.split(':')
while len(range_split) != 2:
  start_end = input("格式错误，请输入开始行和结束行如2:10：")
languages = input("请输入目标语言（可以填中文如德语），多个请用空格隔开：")
lang_split = languages.split(' ')
for lang in lang_split:
  print('开始翻译：', lang)
  # "D:/弹幕白包多语言.xlsx"
  process(
    path=file_path,
    sheet_name=sheet_name,
    column=base_column,
    start=int(range_split[0]),
    end=int(range_split[1]),
    language=lang
  )
print('end...')