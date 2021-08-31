import os
import sys
from tkinter import filedialog

import jinja2
import pandas as pd
import tkinter as tk


window = tk.Tk()
window.withdraw()
file_path = filedialog.askopenfilename(title='選擇來源檔案', filetypes=[("Excel files", ".xlsx .xls")])

if file_path == '':
    sys.exit()


basename = os.path.basename(file_path)
source_filename = os.path.splitext(basename)[0]


def read_excel(file_path: str):
    excel = pd.read_excel(file_path)
    cols = ['收件人姓名', '收件人電話', '收件人地址', '訂購人電話', '訂購人', '訂購人地址', '數量']
    excel = excel[cols]
    excel['收件人電話'].fillna('', inplace=True)
    return excel


def get_parsed_data():
    labels = list()
    for index, row in read_excel(file_path).iterrows():
        if not row.isnull().any():
            for _ in range(row['數量']):
                label = {"sender_tel": row['訂購人電話'],
                         "sender_address": row['訂購人地址'],
                         "contact": row['訂購人'],
                         "receiver_tel": row['收件人電話'],
                         "receiver_address": row['收件人地址'],
                         "receiver_name": row['收件人姓名']
                         }
                labels.append(label)
    return labels


def output(file_name: str):
    templateLoader = jinja2.FileSystemLoader(searchpath="./")
    templateEnv = jinja2.Environment(loader=templateLoader)
    template = templateEnv.get_template("label.html")

    export_path = 'export'
    if not os.path.exists(export_path):
        os.mkdir(export_path)

    with open(export_path + '/' + source_filename + '-' + file_name + '.html', 'a', encoding='utf-8') as file:
        file.write(template.render(
            left_label=left_label,
            right_label=right_label
        ))


left_label = list()
right_label = list()
parsed_data = get_parsed_data()
data_count = len(parsed_data)
i = 0
for label in parsed_data:
    if i >= 8:
        i = 0
        data_count -= 8
        output('訂單標籤')
        left_label.clear()
        right_label.clear()

    if i < 4:
        left_label.append(label)
    elif 3 < i < 8:
        right_label.append(label)
    i += 1

if len(left_label) > 0 or len(right_label):
    output('訂單標籤')
