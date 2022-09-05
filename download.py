import requests
import datetime


def download_cur(from_date, cur='usd', get_links=False):
    match cur:
        case 'usd':
            cur_id = 1235
        case 'cny':
            cur_id = 1375
        case _:
            raise Exception('нет такой валюты в списке')
    from_year, from_month, from_day = map(str, from_date.split('-'))
    today = datetime.date.today()
    day = str(today.day)
    if len(day) < 2:
        day = '0' + day
    month = str(today.month)
    if len(month) < 2:
        month = '0' + month
    year = str(today.year)

    url = f'http://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?Posted=True&so=1&mode=1&VAL_NM_RQ=R0{cur_id}&From={from_day}.{from_month}.{from_year}&To={day}.{month}.{year}&FromDate={from_month}%2F{from_day}%2F{from_year}&ToDate={month}%2F{day}%2F{year}'

    if get_links:
        print(url)
    r = requests.get(url)
    open(f'{cur}.xlsx', 'wb').write(r.content)
