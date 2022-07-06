import requests
import json
import datetime
import xlsxwriter
import requests
import json
import datetime
from tkinter import *
from tkinter import messagebox






def get_token():
    url = "https://api-v3.neuro.net/api/v2/ext/auth"

    payload = {}
    headers = {
        'Authorization': 'Basic ZHNvdG5pa292QGZyb210ZWNoLnJ1OnlaMDZhenlL'
    }
    response = requests.request("POST", url, headers=headers, data=payload)
    data = response.json()
    return data['token']


token = get_token()


def get_name_queue(token, campaign_id):
    agent_uuid_list = ['d1854d56-b39a-4b43-ac22-270763c801e0', '04389590-999b-4f47-a3db-a4c0603ff4dd',
                       '98d066b8-44f1-4c8a-9f61-b337c850fa62', 'ced0a953-ee1b-4298-b183-990dbaa46dde']
    for uuid_uuid_agent in agent_uuid_list:
        url = f"https://api-v3.neuro.net/api/v2/ext/selection?agent_uuid={uuid_uuid_agent}"
        payload = {}
        headers = {
            'Authorization': f'Bearer {token}'
        }
        response = requests.request("GET", url, headers=headers, data=payload)
        data_queue = json.loads(response.text)
        for name in data_queue:
            print(name)
            if campaign_id in name['name']:
                print(name['name'])
                return name['name']



def get_info(token,date_start_search,date_finish_search):
    today = datetime.datetime.today()
    date = datetime.datetime.today() - datetime.timedelta(days=-1)
    workbook = xlsxwriter.Workbook(f'Отказы.xlsx')
    # создаем там "лист"
    worksheet = workbook.add_worksheet()
    # в ячейку A1 пишем текст
    worksheet.write(f'A1', 'Номер абонента')
    worksheet.write(f'B1', 'Регион')
    worksheet.write(f'C1', 'Номер и название компании')
    worksheet.write(f'D1', 'Дата отказа')
    worksheet.write(f'E1', 'Опция')
    change_count = 1
    agent_uuid_list = ['d1854d56-b39a-4b43-ac22-270763c801e0', '04389590-999b-4f47-a3db-a4c0603ff4dd',
                       '98d066b8-44f1-4c8a-9f61-b337c850fa62', 'ced0a953-ee1b-4298-b183-990dbaa46dde']
    for agent_uuid in agent_uuid_list:
        print(f'Ищем в агенте {agent_uuid}')
        payload = json.dumps({
          "agent_uuid": f"{agent_uuid}",
          #"start": f"2022-06-29 21:00:33.0",
          #"end": "2022-06-30 20:59:33.0",
          "start": f"2022-{date_start_search} 00:01:00.0",
          "end": f"2022-{date_finish_search} 23:59:00.0",
          "limit": 100000,
          "offset": 0
        })
        headers = {
          'Authorization': f'Bearer {token}',
          'Content-Type': 'application/json'
        }
        url = "https://api-v3.neuro.net/api/v2/ext/statistic/dialog-report"
        response = requests.request("POST", url, headers=headers, data=payload)
        data = json.loads(response.text)
        numbers_in_file = []
        for dialog_info in data['result']:
            #print(f'Словарь {dialog_info}')
            if 'status_scheme' in dialog_info and dialog_info['status_scheme'] == 'Отмена перехода':
                print('Нашлась отмена перехода')
                if dialog_info['msisdn'] in numbers_in_file:
                    continue
                numbers_in_file.append(dialog_info['msisdn'])
                change_count += 1
                name_queue = get_name_queue(token, dialog_info['campaign_id'])
                print(name_queue)
                #lambda name: name_queue.split('t2_upload_')[1][:-5] if 'RELOAD' in name_queue else name_queue.split('')('t2_upload_')[1][:-12])
                if 'RELOAD' not in name_queue:
                    worksheet.write(f'C{str(change_count)}', name_queue.split('t2_upload_')[1][:-5])
                else:
                    worksheet.write(f'C{str(change_count)}', name_queue.split('t2_upload_')[1][:-12])
                worksheet.write(f'A{str(change_count)}', dialog_info['msisdn'])
                worksheet.write(f'B{str(change_count)}', dialog_info['branch_name'])
                worksheet.write(f'D{str(change_count)}', dialog_info['CallDateTime'].replace('-','.')[:-15])
                worksheet.write(f'E{str(change_count)}', 'Отменить смену ТП (Отказ от предложения (после согласия))')
                print(f'Найден и записан номер {dialog_info["msisdn"]}, {dialog_info["branch_name"]}, {dialog_info["CallDateTime"]}. Записываем в строку {str(change_count)}')
    print(f'Закончили с агентом {agent_uuid}')
    workbook.close()


def start_foo():
    get_info(token,name.get(),surname.get())
    root.destroy()


root = Tk()
root.title("Питомец для поиска отмен перехода")
root['background']='#007ba7'
name = StringVar()
surname = StringVar()

name_label = Label(text="Дата начала:")
surname_label = Label(text="Дата конца:")

name_label.grid(row=0, column=0, sticky="w")
surname_label.grid(row=1, column=0, sticky="w")

name_entry = Entry(textvariable=name)
surname_entry = Entry(textvariable=surname)

name_entry.grid(row=0, column=1, padx=5, pady=5)
surname_entry.grid(row=1, column=1, padx=5, pady=5)

message_button = Button(text="Сформировать отчёт", command=start_foo)
message_button.grid(row=2, column=1, padx=5, pady=5, sticky="e")

root.mainloop()