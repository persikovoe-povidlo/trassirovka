import pandas as pd
import time
import datetime
from pyexcelerate import Workbook


def main():
    total_time = time.time()

    file = 'Сделки_01072022-03082022_03-08-2022_174758.xls'
    eod_prices_file = 'Pozitsii_Po_Tsb_01072022-04082022_04-08-2022_155347.xls'
    leftovers_file = 'Pozitsii_Po_Tsb_30062022-30062022_09-08-2022_103732.xls'
    exch_rates_usd_file = 'usd.xlsx'
    exch_rates_cny_file = 'cny.xlsx'
    positions = []

    start_time = time.time()
    print('reading files...')

    x = pd.concat(
        [pd.read_excel(file, sheet_name=0, usecols=[3, 8, 11, 12, 13, 20, 21, 23, 25, 43, 44, 55], header=None),
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

    leftovers_list = pd.read_excel(leftovers_file, sheet_name=0, usecols=[0, 8, 11, 13, 17, 21, 26, 28, 29],
                                   header=None)
    leftovers_df = pd.DataFrame()
    for i in range(leftovers_list[0].size):
        date = str(leftovers_list[0][i]).split()[0]
        stock_name = leftovers_list[8][i]
        stock_amount = leftovers_list[13][i]
        if date == needed_date and stock_name in positions.keys() and stock_amount:
            stock_type = leftovers_list[11][i]
            stock_price = leftovers_list[17][i]
            currency_name = leftovers_list[21][i]
            aci = leftovers_list[26][i]
            currency_price_rub = round(leftovers_list[28][i] / stock_price / stock_amount, 9)
            denomination = leftovers_list[29][i]
            new_df = pd.DataFrame(
                [['', '', '', '', '', '', '', '', stock_name, '', '', denomination, stock_type, date, '', '', '',
                  '', '', '', stock_amount, stock_price, '', aci, '', currency_name, '', '', '', '', '', '', '', '',
                  '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                  currency_price_rub]])
            leftovers_df = pd.concat([leftovers_df, new_df])
            if date not in eod_price_dict:
                eod_price_dict[date] = {}
            eod_price_dict[date][stock_name] = stock_price * currency_price_rub
            eod_price_dict[date]['РУБ'] = 1
    x = pd.concat([leftovers_df, x], ignore_index=True)

    exch_rates_usd = pd.read_excel(exch_rates_usd_file, sheet_name=0, usecols=[1, 2], header=0)
    usd_prices = exch_rates_usd.to_dict('records')

    exch_rates_cny = pd.read_excel(exch_rates_cny_file, sheet_name=0, usecols=[0, 1, 2], header=0)
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

    for i in range(n):
        if i >= n / 10 * k:
            print('{} % done'.format(k * 10))
            k += 1
        stock_name = str(x[8][i])
        stock_amount = x[20][i]
        currency_price_rub = round(x[55][i], 9)
        stock_price_rub = x[21][i] * currency_price_rub
        currency_name = str(x[25][i])
        currency_amount = round(x[21][i] * -stock_amount, 9)
        if x[12][i] in ['Еврооблигации', 'Облигиция', 'ОФЗ']:
            denomination = x[11][i] / 100
            stock_price_rub = stock_price_rub * denomination
            currency_amount = currency_amount * denomination

        date = str(x[13][i]).split()[0]
        next_date = str(x[13][i + 1]).split()[0]

        aci = round(x[23][i], 9)

        repo_cell = str(x[43][i])

        if repo_cell == 'РЕПО 1 часть':
            repos[x[3][i]] = currency_amount
            realized_stock = None
            realized_cur = None
            stock_fifo_amount = None
            currency_fifo_amount = None
            stock_fifo = None
            currency_fifo = None
        elif repo_cell == 'РЕПО 2 часть':
            realized_stock = currency_amount + repos[x[44][i]]
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


if __name__ == '__main__':
    main()
