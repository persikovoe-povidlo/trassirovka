import pandas as pd
import time
import datetime
from pyexcelerate import Workbook
from download import download_cur
import warnings
from os.path import exists
from os import remove, getlogin
import webbrowser


def main():
    total_time = time.time()

    file = 'Сделки_01072022-03082022_03-08-2022_174758.xls'
    eod_prices_file = 'Pozitsii_Po_Tsb_01072022-04082022_04-08-2022_155347.xls'
    leftovers_file = 'Pozitsii_Po_Tsb_30062022-30062022_09-08-2022_103732.xls'
    last_day_files_downloaded_file = 'last_day_files_downloaded.txt'
    exch_rates_usd_file = 'usd.xlsx'
    exch_rates_cny_file = 'cny.xlsx'
    positions = []

    start_time = time.time()
    print('reading files...')

    if not exists(last_day_files_downloaded_file):
        open(last_day_files_downloaded_file, 'w')

    download_currencies(last_day_files_downloaded_file)

    x = pd.concat(
        [pd.read_excel(file, sheet_name=0, usecols=[3, 8, 13, 20, 25, 29, 30, 43, 44, 55], header=None),
         pd.DataFrame([[]])], ignore_index=True).drop([0])
    x[13] = pd.to_datetime(x[13])
    x = x.sort_values(by=[13, 43], ignore_index=True)

    needed_date = str(x[13][1] - datetime.timedelta(days=1)).split()[0]

    eod_price_list = pd.read_excel(eod_prices_file, sheet_name=0, usecols=[0, 8, 13, 24], header=None)
    eod_price_dict = {}
    stock_positions = x[8].values.tolist()
    cur_positions = x[25].values.tolist()
    stock_positions.pop(0)
    stock_positions.pop(-1)
    cur_positions.pop(0)
    cur_positions.pop(-1)
    for e in stock_positions + cur_positions:
        if e not in positions:
            positions.append(e)
    positions = {e: 0 for e in positions}
    queues = {e: [] for e in positions}

    leftovers_dict = pd.read_excel(leftovers_file, sheet_name=0, usecols=[0, 8, 13, 17, 21, 24, 26],
                                   header=None).to_dict('records')

    if needed_date not in eod_price_dict:
        eod_price_dict[needed_date] = {}

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        exch_rates_usd = pd.read_excel(exch_rates_usd_file, sheet_name=0, usecols=[1, 2], header=0, engine='openpyxl')
    usd_prices = exch_rates_usd.to_dict('records')

    with warnings.catch_warnings(record=True):
        warnings.simplefilter("always")
        exch_rates_cny = pd.read_excel(exch_rates_cny_file, sheet_name=0, usecols=[0, 1, 2], header=0,
                                       engine='openpyxl')
    cny_prices = exch_rates_cny.to_dict('records')

    for i in range(1, eod_price_list[0].size):
        date = str(eod_price_list[0][i]).split()[0]
        if date not in eod_price_dict:
            eod_price_dict[date] = {}
        amount = eod_price_list[13][i]
        if amount:
            eod_price_dict[date][eod_price_list[8][i]] = eod_price_list[24][i] / eod_price_list[13][i]
            eod_price_dict[date]['РУБ'] = 1

    for row in usd_prices:
        date = str(row['data']).split()[0]
        if date in eod_price_dict.keys():
            eod_price_dict[date]['USD'] = row['curs']

    for row in cny_prices:
        date = str(row['data']).split()[0]
        if date in eod_price_dict.keys():
            eod_price_dict[date]['CNY'] = row['curs'] / row['nominal']

    leftovers_list = []

    for row in leftovers_dict:
        date = str(row[0]).split()[0]
        stock_name = row[8]
        stock_amount = row[13]
        if date == needed_date and stock_name in positions.keys() and stock_amount:
            total_stock_price = row[24]
            stock_price_rub = round(total_stock_price / abs(stock_amount), 9)
            currency_name = row[21]
            aci = round(row[26] / abs(stock_amount), 9)
            eod_price_dict[date][stock_name] = stock_price_rub
            eod_price_dict[date]['РУБ'] = 1
            currency_price_rub = eod_price_dict[date][currency_name]
            leftovers_list.append(['', '', '', '', '', '', '', '', stock_name, '', '', '', '', date, '', '', '',
                                   '', '', '', stock_amount, '', '', '', '', currency_name, '', '', '',
                                   -total_stock_price, -aci, '', '', '',
                                   '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                   currency_price_rub])
    leftovers_df = pd.DataFrame(leftovers_list)
    x = pd.concat([leftovers_df, x], ignore_index=True)

    eod_price = dict(positions)
    imp_sum = 0
    not_imp_last_day = 0
    repos = {}

    table = [['', '', '', '', '', 'позиции', '', '', '', *list(' ' * len(positions)), 'цена на конец дня'],
             ['', '', 'количество', 'цена руб', 'нкд', *positions, '', 'кол-во для расчёта финреза', 'цена ФИФО',
              'реал', *eod_price, 'накопл финрез', 'реал накопл', 'нереал накопл', 'нереализ дневной']]

    print("%s seconds" % round(time.time() - start_time, 2))
    start_time = time.time()
    print('main script running...')

    n = x[8].size - 1
    k = 1

    x = x.to_dict('records')
    for i, row in enumerate(x[:-1]):
        if i >= n / 10 * k:
            print('{} % done'.format(k * 10))
            k += 1
        stock_name = str(row[8])
        stock_amount = row[20]
        currency_price_rub = round(row[55], 9)
        stock_price_rub = row[29] / -stock_amount
        currency_name = str(row[25])
        currency_amount = round(stock_price_rub / currency_price_rub * -stock_amount, 9)

        date = str(row[13]).split()[0]
        next_date = str(x[i + 1][13]).split()[0]
        aci = round(abs(row[30] / stock_amount), 9)

        repo_cell = str(row[43])

        if repo_cell == 'РЕПО 1 часть':
            repos[row[3]] = currency_amount
            realized_stock = None
            realized_cur = None
            stock_fifo_amount = None
            currency_fifo_amount = None
            stock_fifo = None
            currency_fifo = None
        elif repo_cell == 'РЕПО 2 часть':
            realized_stock = currency_amount + repos[row[44]]
            realized_cur = None
            stock_fifo_amount = None
            currency_fifo_amount = None
            stock_fifo = None
            currency_fifo = None
        else:
            stock_fifo_amount = get_fifo_amount(stock_name, stock_amount, positions)
            currency_fifo_amount = get_fifo_amount(currency_name, currency_amount, positions)
            stock_fifo = get_fifo(stock_fifo_amount, stock_name, stock_amount, stock_price_rub, queues, aci)
            currency_fifo = get_fifo(currency_fifo_amount, currency_name, currency_amount, currency_price_rub, queues,
                                     0)
            realized_stock = get_realized(stock_fifo_amount, stock_fifo, stock_price_rub, aci)
            realized_cur = get_realized(currency_fifo_amount, currency_fifo, currency_price_rub, 0)

        positions[stock_name] += stock_amount
        positions[currency_name] += currency_amount

        if stock_fifo_amount:
            stock_fifo_amount = round(stock_fifo_amount, 9)
        if currency_fifo_amount:
            currency_fifo_amount = round(currency_fifo_amount, 9)

        if realized_stock:
            realized_stock = round(realized_stock, 9)
            imp_sum += realized_stock
        if realized_cur:
            realized_cur = round(realized_cur, 9)
            imp_sum += realized_cur

        if date != next_date:
            acc_fifo_amount = 0
            for e in eod_price:
                if e in eod_price_dict[date]:
                    eod_price[e] = eod_price_dict[date][e]
                acc_fifo_amount += positions[e] * eod_price[e]
            acc_fifo_amount = round(acc_fifo_amount, 9)
            not_imp = acc_fifo_amount - imp_sum
            not_imp_day = not_imp - not_imp_last_day
            not_imp_last_day = not_imp
            table.append([date, stock_name, stock_amount, stock_price_rub, aci, '', *list(' ' * len(positions)),
                          stock_fifo_amount, stock_fifo, realized_stock])
            table.append(
                [date, currency_name, currency_amount, currency_price_rub, aci, '', *list(' ' * len(positions)),
                 currency_fifo_amount, currency_fifo, realized_cur])
            table.append(['', '', '', '', '', *[positions[e] for e in positions], '', '', '', '',
                          *[eod_price[e] for e in eod_price], acc_fifo_amount, imp_sum, not_imp, not_imp_day])
        else:
            table.append([date, stock_name, stock_amount, stock_price_rub, aci, '', *list(' ' * len(positions)),
                          stock_fifo_amount, stock_fifo, realized_stock, ])
            table.append(
                [date, currency_name, currency_amount, currency_price_rub, aci, '', *list(' ' * len(positions)),
                 currency_fifo_amount, currency_fifo, realized_cur])
            table.append(['', '', '', '', '', *[positions[e] for e in positions]])

    df = pd.DataFrame(table)

    print('{} % done'.format(k * 10))
    print("%s seconds" % round(time.time() - start_time, 2))
    start_time = time.time()
    print('writing to file...')

    save_to_csv(df, 'out.csv')
    # save_to_xlsx(df, 'out.xlsx')

    print("%s seconds" % round(time.time() - start_time, 2))
    print('--------------------')
    print("%s seconds total" % round(time.time() - total_time, 2))


def get_realized(fifo_amount, fifo, price, aci):
    if fifo_amount:
        return fifo_amount * (fifo - price - aci)


def get_fifo_amount(name, amount, positions):
    fifo_amount = None
    if positions[name] * (positions[name] + amount) > 0:
        if name != 'РУБ' and abs(positions[name] + amount) < abs(positions[name]):
            fifo_amount = amount
    else:
        if positions[name] and name != 'РУБ':
            fifo_amount = -positions[name]
    return fifo_amount


def get_fifo(fifo_amount, name, amount, price, queues, aci):
    fifo = None
    if fifo_amount:
        fifo = 0
        sum_amount = 0
        if abs(queues[name][0][0]) < abs(amount):
            for e in queues[name]:
                if abs(amount) >= sum_amount + abs(e[0]):
                    fifo += abs(e[0]) * (e[2] + e[1])
                elif abs(amount) > abs(sum_amount):
                    fifo += (abs(amount) - sum_amount) * (e[2] + e[1])
                sum_amount += abs(e[0])
            fifo /= abs(fifo_amount)
        else:
            fifo = queues[name][0][1] + queues[name][0][2]
        fifo = round(fifo, 9)

    if queues[name]:
        if queues[name][0][0] * amount <= 0:
            queues[name].insert(0, [amount, price, aci])
            while len(queues[name]) > 1:
                if abs(queues[name][0][0]) < abs(queues[name][1][0]):
                    queues[name][0] = [queues[name][0][0] + queues[name][1][0], queues[name][1][1], queues[name][1][2]]
                else:
                    queues[name][0] = [queues[name][0][0] + queues[name][1][0], queues[name][0][1], queues[name][0][2]]
                queues[name].pop(1)
        else:
            queues[name].append([amount, price, aci])
    else:
        queues[name].append([amount, price, aci])
    return fifo


def save_to_csv(df, filename):
    df.to_csv(filename, index=False, header=False, encoding='utf-8-sig')


def save_to_xlsx(df, filename, sheetname='1'):
    values = [df.columns] + list(df.values)
    wb = Workbook()
    wb.new_sheet(sheetname, data=values)
    wb.save(filename)


def download_currencies(last_day_files_downloaded_file):
    with open(last_day_files_downloaded_file, 'r+') as f:
        text = f.readline()
        date = str(datetime.date.today())
        if text != date:
            f.seek(0)
            f.write(str(date))
            f.truncate()
            for cur in ['usd', 'cny']:
                download_cur('2022-01-01', cur)
                print(f'downloaded {cur}.xlsx')


def bypass_site_protection():
    webbrowser.open(
        'http://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?Posted=True&so=1&mode=1&VAL_NM_RQ=R01235&From=01.01.2022&To=01.01.2022&FromDate=01%2F01%2F2022&ToDate=01%2F01%2F2022',
        new=1)
    time.sleep(2)

    remove('C:\\Users\\' + getlogin() + '\\Downloads\\RC_F01_01_2022_T01_01_2022.xlsx')


if __name__ == '__main__':
    main()
