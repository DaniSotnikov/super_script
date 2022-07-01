import requests
import json
import datetime
import xlsxwriter
import requests
import json
import datetime


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

def get_name_queue(token,campaign_id):
    print('Зашли в функцию')
    agent_uuid_list = ['d1854d56-b39a-4b43-ac22-270763c801e0', '04389590-999b-4f47-a3db-a4c0603ff4dd',
                       '98d066b8-44f1-4c8a-9f61-b337c850fa62', 'ced0a953-ee1b-4298-b183-990dbaa46dde']
    for agent_uuid in agent_uuid_list:
        print('Взяли новый uuid')
        url = f"https://api-v3.neuro.net/api/v2/ext/selection?agent_uuid={agent_uuid}"
        payload = {}
        headers = {
            'Authorization': f'Bearer {token}'
        }
        response = requests.request("GET", url, headers=headers, data=payload)
        data_queue = json.loads(response.text)
        for name in data_queue:
            if name['name'].find(campaign_id):
                continue
                #return name['name']
            else:
                continue
    print('Вообще не найдено')

def get_info(token):
    today = datetime.datetime.today()
    date = datetime.datetime.today() - datetime.timedelta(days=-1)
    print(date)
    workbook = xlsxwriter.Workbook(f'Отказы.xlsx')
    # создаем там "лист"
    worksheet = workbook.add_worksheet()
    # в ячейку A1 пишем текст
    worksheet.write(f'A1', 'Номер абонента')
    worksheet.write(f'B1', 'Номер и название компании')
    worksheet.write(f'C1', 'Дата отказа')
    worksheet.write(f'D1', 'Опция')
    change_count = 1
    test_data = {'result':[
                     {'CallDateTime': '2022-06-30 22:06:28.993090','campaign_id': '265843','branch_name': 'Саратов New', 'msisdn': '+795242101251',
                      'status_scheme': 'Отмена перехода'},
                 {'CallDateTime': '2022-06-30 22:06:28.993090','campaign_id': '265843','branch_name': 'Саратов New', 'msisdn': '+7952213101251',
                  'status_scheme': 'Отмена перехода'}]}
    agent_uuid_list = ['d1854d56-b39a-4b43-ac22-270763c801e0', '04389590-999b-4f47-a3db-a4c0603ff4dd',
                       '98d066b8-44f1-4c8a-9f61-b337c850fa62', 'ced0a953-ee1b-4298-b183-990dbaa46dde']
    for agent_uuid in agent_uuid_list:
        print(f'Ищем Отмену перехода в {agent_uuid}')
        # payload = json.dumps({
        #   "agent_uuid": f"{agent_uuid}",
        #   "start": "2022-06-29 21:00:33.0",
        #   "end": "2022-06-30 20:59:33.0",
        #   "limit": 1000,
        #   "offset": 0
        # })
        # headers = {
        #   'Authorization': f'Bearer {token}',
        #   'Content-Type': 'application/json'
        # }
        #
        # response = requests.request("POST", url, headers=headers, data=payload)
        # data = json.loads(response.text)
        for dialog_info in test_data['result']:
            if 'status_scheme' in dialog_info and dialog_info['status_scheme'] == 'Отмена перехода':
                change_count +=1
                worksheet.write(f'A{str(change_count)}', dialog_info['msisdn'])
                worksheet.write(f'B{str(change_count)}', dialog_info['branch_name'])
                worksheet.write(f'C{str(change_count)}', get_name_queue(token,dialog_info['campaign_id'],agent_uuid))
                worksheet.write(f'D{str(change_count)}', 'Отменить смену ТП (Отказ от предложения (после согласия))')
            else:
                continue
        print("Закончили поиск")
    workbook.close()

#get_info(token)
#get_info(token)
get_name_queue(token,'246903')