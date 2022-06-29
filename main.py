import pandas as pd
import numpy as np
from datetime import datetime
import threading
from tkinter import *
from tkinter import filedialog as fd
import tkinter as tk
import tkinter.font as tkFont
from PIL import ImageTk, Image
name_campagin = []



def generate_analytics_for_compaing(file_original_name: str):
    print('Yfxfkj aeyrwbb')
    df = pd.read_excel(file_original_name)
    print('Читаем файл')
    print('123')
    for campaign_id in pd.unique(df['campaign_id']):
        if campaign_id == 'None':
            continue
        name_campagin.append(campaign_id)
        threads = threading.Thread(target=starting_threading, args=(campaign_id, df))
        threads.start()


def starting_threading(campaign_id, df):
    selection_by_campaign_id = df[df['campaign_id'] == campaign_id]
    print(f'Создаём файл по выборке {campaign_id}')
    name_new_file = f'Fromtech_{campaign_id}.xlsx'
    selection_by_campaign_id.to_excel(name_new_file)
    drop_coloumns(name_new_file, pd.read_excel(name_new_file))


def drop_coloumns(name_file, df):
    with pd.ExcelWriter(name_file) as writer:
        for name_columns in df.columns:
            if name_columns != 'msisdn' and name_columns != 'CallDateTime' and name_columns != 'status_scheme':
                df.drop(columns=name_columns, axis=1, inplace=True)
        print(f'Удалили столбцы из {name_file}')
        new_df = df.replace(np.nan, 'Недозвон', regex=True)
        print('Заменили пустые статусы на наедозвон')
        count_of_rows = new_df.count()
        count_of_rows_call_date = count_of_rows['CallDateTime']
        for i in range(int(count_of_rows_call_date)):
            if new_df.iloc[i]['CallDateTime'] == 'Недозвон':
                new_df.loc[[i], 'CallDateTime'] = new_df.iloc[i - 1]['CallDateTime'][:-7]
            else:
                new_df.loc[[i], 'CallDateTime'] = new_df.iloc[i]['CallDateTime'][:-7]
            if str(new_df.iloc[i]['msisdn'])[0] == '9':
                new_df.loc[[i], 'msisdn'] = '7' + str(new_df.iloc[i]['msisdn'])
        new_df.rename(columns={'msisdn': 'MSISDN', 'status_scheme': 'Результат звонка'}, inplace=True)
        print('Записываем в файл')
        new_df.to_excel(writer, index=False)

def insert_file():
    file_name = fd.askopenfilename()
    print(file_name)
    generate_analytics_for_compaing(file_name)

root = Tk()
root.title("Графическая программа на Python")
a = root.geometry()
fontStyle = tkFont.Font(family="Lucida Grande", size=20)
b1 = Button(text="Сформировать отчёт", height=30,width=40, command=insert_file, bg='#ffc0cb',compound=tk.CENTER)
b1.grid(row=500, column=100, ipadx=30, ipady=6, padx=600, pady=130)

root.mainloop()

