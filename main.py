import datetime
import os
from tkinter import *
from tkinter import filedialog as fd
import tkinter as tk
import tkinter.font as tkFont
import pandas as pd
import numpy as np
print('Ожидаем ввод файла')

name_campagin = []
from multiprocessing import Process


def generate_analytics_for_compaing(file_original_name: str):
    print('Читаем файл')
    df = pd.read_excel(file_original_name)
    for campaign_id in pd.unique(df['campaign_id']):
        if campaign_id == 'None':
            continue
        name_campagin.append(campaign_id)
        starting_threading(campaign_id, df)


def starting_threading(campaign_id, df):
    selection_by_campaign_id = df[df['campaign_id'] == campaign_id]
    print(f'Создаём файл по выборке {campaign_id}')
    homeDir = os.path.expanduser('~') + r'\Desktop'
    name_new_file = f'Fromtech_000{campaign_id}.xlsx'
    new_selection_withour_inbound = selection_by_campaign_id[selection_by_campaign_id['call_direction'] == 'outbound']
    new_selection_withour_inbound.to_excel(os.path.join(homeDir,name_new_file))
    df = pd.read_excel(os.path.join(homeDir,name_new_file))
    with pd.ExcelWriter(os.path.join(homeDir,name_new_file)) as writer:
        for name_columns in df.columns:
            if name_columns != 'route_code' and name_columns != 'CallDateTime' and name_columns != 'status_scheme':
                df.drop(columns=name_columns, axis=1, inplace=True)
        print(f'Удалили столбцы из {name_new_file}')
        new_df = df.replace(np.nan, 'Недозвон', regex=True)
        print('Заменили пустые статусы на Недозвон')
        count_of_rows = new_df.count()
        count_of_rows_route_code = count_of_rows['route_code']
        for i in range(int(count_of_rows_route_code )):
            if new_df.iloc[i]['status_scheme'] in ('Исходящий запрещен, ранее состоялся диалог', 'Исходящий запрещен, был входящий перезвон'):
                new_df.loc[[i], 'status_scheme'] = ''
                new_df.loc[[i], 'route_code'] = ''
                new_df.loc[[i], 'CallDateTime'] = ''
            else:
                if new_df.iloc[i]['CallDateTime'] == 'Недозвон':
                    #new_df.loc[[i], 'CallDateTime']
                    try:
                        date_str = datetime.datetime.strptime(str(new_df.iloc[i - 1]['CallDateTime']).replace('-', '.'), '%Y.%m.%d %H:%M:%S.%f')
                        new_df.loc[[i], 'CallDateTime'] = date_str.strftime('%d.%m.%Y %H:%M')
                    except ValueError:
                        new_df.loc[[i], 'CallDateTime'] = str(new_df.iloc[i - 1]['CallDateTime'])
                else:
                    new_df.loc[[i], 'CallDateTime'] = datetime.datetime.strptime(str(new_df.iloc[i]['CallDateTime']).replace('-', '.'), '%Y.%m.%d %H:%M:%S.%f').strftime('%d.%m.%Y %H:%M')
                if str(new_df.iloc[i]['route_code'])[0] == '9':
                    new_df.loc[[i], 'route_code'] = '7' + str(new_df.iloc[i]['route_code'][1:])
        #new_new_df = new_df.set_index('status_scheme')
        #new_new_df = new_new_df.drop(['Исходящий запрещен, ранее состоялся диалог'],axis=0)
        filter = new_df['route_code'] != ''
        new_df = new_df[filter]
        new_df.rename(columns={'route_code': 'MSISDN', 'status_scheme': 'Результат звонка'}, inplace=True)
        print('Редактируем и сохраняем файл')
        new_df.to_excel(writer,index=False)

def insert_file():
    file_name = fd.askopenfilename()
    print(file_name)
    generate_analytics_for_compaing(file_name)
    root.destroy()


if __name__ == '__main__':
    root = Tk()
    root.title("Графическая программа на Python")
    a = root.geometry('140x150')
    root.resizable(False,False)
    fontStyle = tkFont.Font(family="Lucida Grande", size=10)
    b1 = Button(text="Сформировать отчёт", height=10, width=10, command=insert_file, bg='#ffc0cb', compound=tk.CENTER)
    b1.grid(row=500, column=100, ipadx=30, ipady=6, padx=0, pady=0)
    root.mainloop()




