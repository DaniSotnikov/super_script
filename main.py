import pandas as pd
import numpy as np
from datetime import datetime
name_campagin = []


def generate_analytics_for_compaing(file_original_name: str):
   df = pd.read_excel(f'{file_original_name}.xlsx')
   for campaign_id in pd.unique(df['campaign_id']):
       name_campagin.append(campaign_id)
       selection_by_campaign_id = df[df['campaign_id'] == campaign_id]
       print(f'Создаём файл по выборке {campaign_id}')
       name_new_file = f'Fromtech_{campaign_id}.xlsx'
       selection_by_campaign_id.to_excel(name_new_file)
       drop_coloumns(name_new_file,pd.read_excel(name_new_file))


def drop_coloumns(name_file,df):
    with pd.ExcelWriter(name_file) as writer:
        for name_coloumns in df.columns:
            if name_coloumns != 'msisdn' and name_coloumns != 'CallDateTime' and name_coloumns != 'status_scheme':
                df.drop(columns=name_coloumns, axis=1, inplace=True)
        print(f'Удалили столбцы из {name_file}')
        k = df.replace(np.nan,'Недозвон',regex=True)
        print('Заменили пустые статусы на наедозвон')
        count_of_rows = k.count()
        count_of_rows_call_date = count_of_rows['CallDateTime']
        for i in range(int(count_of_rows_call_date)):
            if k.iloc[i]['CallDateTime'] == 'Недозвон':
                k.loc[[i],'CallDateTime'] = k.iloc[i-1]['CallDateTime']
        k.rename(columns={'msisdn': 'MSISDN','status_scheme': 'Результат звонка'}, inplace=True)
        k.to_excel(writer)
time_now = datetime.now()
generate_analytics_for_compaing('Test')
time_after = datetime.now()
print(time_now - time_after)

#k['CallDateTime'] = np.where((k['CallDateTime'] == 'Недозвон'), 'Заменили время', k['CallDateTime'])

