import requests
import datetime


# Download manualy if downloaded files are broken. Then should start working fine
# http://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?Posted=True&so=1&mode=1&VAL_NM_RQ=R01235&From=01.01.2022&To=13.08.2022&FromDate=01%2F01%2F2022&ToDate=08%2F13%2F2022
def download_cur(from_day='01', from_month='01', from_year='2022', cur='usd'):
    match cur:
        case 'usd':
            cur_id = 1235
        case 'cny':
            cur_id = 1375
        case _:
            raise Exception('нет такой валюты в списке')
    today = datetime.date.today()
    day = str(today.day)
    if len(day) < 2:
        day = '0' + day
    month = str(today.month)
    if len(month) < 2:
        month = '0' + month
    year = str(today.year)

    url = f'http://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?Posted=True&so=1&mode=1&VAL_NM_RQ=R0{cur_id}&From={from_day}.{from_month}.{from_year}&To={day}.{month}.{year}&FromDate={from_month}%2F{from_day}%2F{from_year}&ToDate={month}%2F{day}%2F{year}'

    r = requests.get(url)
    open(f'{cur}.xlsx', 'wb').write(r.content)


download_cur('01', '01', '2022', 'usd')
download_cur('01', '01', '2022', 'cny')
