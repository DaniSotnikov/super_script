import pandas as pd
import numpy as np
from datetime import datetime
name_campagin = []


def generate_analytics_for_compaing(file_original_name: str):
   df = pd.read_excel(f'{file_original_name}.xlsx')
   print('Читаем файл')
   for campaign_id in pd.unique(df['campaign_id']):
       if campaign_id == 'None':
           continue
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
        new_df = df.replace(np.nan,'Недозвон',regex=True)
        print('Заменили пустые статусы на наедозвон')
        count_of_rows = new_df.count()
        count_of_rows_call_date = count_of_rows['CallDateTime']
        for i in range(int(count_of_rows_call_date)):
            if new_df.iloc[i]['CallDateTime'] == 'Недозвон':
                new_df.loc[[i],'CallDateTime'] = new_df.iloc[i-1]['CallDateTime'][:-7]
            else:
                new_df.loc[[i], 'CallDateTime'] = new_df.iloc[i]['CallDateTime'][:-7]
            if str(new_df.iloc[i]['msisdn'])[0] == '9':
                new_df.loc[[i], 'msisdn'] = '7' + str(new_df.iloc[i]['msisdn'])
        new_df.rename(columns={'msisdn': 'MSISDN','status_scheme': 'Результат звонка'}, inplace=True)
        new_df.to_excel(writer,index=False)
time_now = datetime.now()
generate_analytics_for_compaing('T2_SUPER_ONLINE_DISCOUNT_from_16_06_2022_00_00_00_to_29_06_2022')
time_after = datetime.now()
print(time_after - time_now)

