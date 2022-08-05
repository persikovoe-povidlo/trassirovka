import pandas as pd
import time


def main():
    total_time = time.time()

    file = 'Сделки_01072022-03082022_03-08-2022_174758.xlsx'
    eod_prices_file = 'Pozitsii_Po_Tsb_01072022-04082022_04-08-2022_155347.xlsx'
    sheet = '1'
    positions = []

    start_time = time.time()
    print('reading file...')

    x = pd.concat(
        [pd.read_excel(file, sheet_name=sheet, usecols=[8, 11, 12, 13, 20, 21, 23, 25, 29, 43, 55], header=None),
         pd.DataFrame([[]])], ignore_index=True)
    eod_price_list = pd.read_excel(eod_prices_file, sheet_name=sheet, usecols=[0, 8, 24], header=None)
    eod_price_dict = {}
    for i in range(1, eod_price_list[0].size):
        date = str(eod_price_list[0][i]).split()[0]
        if date not in eod_price_dict:
            eod_price_dict[date] = {}
        eod_price_dict[date][eod_price_list[8][i]] = eod_price_list[24][i]

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
    eod_price = dict(positions)
    imp_sum = 0
    not_imp_last_day = 0

    df = pd.DataFrame()

    print("%s seconds" % round(time.time() - start_time, 2))
    start_time = time.time()
    print('main script running...')

    for i in range(1, x[8].size - 1):
        stock_name = str(x[8][i])
        stock_amount = x[20][i]
        currency_price_rub = round(x[55][i], 9)
        stock_price_rub = x[21][i] * currency_price_rub
        currency_name = str(x[25][i])
        currency_amount = round(x[21][i] * stock_amount, 9)
        if x[12][i] in ['Еврооблигации', 'Облигиция', 'ОФЗ']:
            stock_price_rub = stock_price_rub * x[11][i] / 100
            currency_amount = currency_amount * x[11][i] / 100

        date = str(x[13][i]).split()[0]
        next_date = str(x[13][i + 1]).split()[0]

        aci = round(x[23][i], 9)

        if str(x[43][i]) == 'nan':
            repo = False
        else:
            repo = True

        stock_fifo_amount = get_fifo_amount(stock_name, stock_amount, positions, repo)
        currency_fifo_amount = get_fifo_amount(currency_name, currency_amount, positions, repo)

        positions[stock_name] += stock_amount
        positions[currency_name] += currency_amount

        stock_fifo = get_fifo(stock_fifo_amount, stock_name, stock_amount, stock_price_rub, queues, aci)
        currency_fifo = get_fifo(currency_fifo_amount, currency_name, currency_amount, currency_price_rub, queues, 0)

        if stock_fifo_amount:
            stock_fifo_amount = round(stock_fifo_amount, 9)
        if currency_fifo_amount:
            currency_fifo_amount = round(currency_fifo_amount, 9)

        realized_stock = get_realized(stock_fifo_amount, stock_fifo, stock_price_rub, aci)
        realized_cur = get_realized(currency_fifo_amount, currency_fifo, currency_price_rub, 0)
        if realized_stock:
            realized_stock = round(realized_stock, 9)
            imp_sum += realized_stock
        if realized_cur:
            realized_cur = round(realized_cur, 9)
            imp_sum += realized_cur

        if date != next_date:
            acc_fifo_amount = positions['РУБ']
            for e in eod_price:
                if e in eod_price_dict[date]:
                    eod_price[e] = eod_price_dict[date][e]
                acc_fifo_amount += positions[e] * eod_price[e]
            acc_fifo_amount = round(acc_fifo_amount, 9)
            not_imp = acc_fifo_amount - imp_sum

            not_imp_day = not_imp - not_imp_last_day
            not_imp_last_day = not_imp
            new_df = pd.DataFrame(
                [[date, stock_name, stock_amount, stock_price_rub, '', '', *list(' ' * len(positions)),
                  stock_fifo_amount, stock_fifo, realized_stock],
                 [date, currency_name, currency_amount, currency_price_rub, '', '', *list(' ' * len(positions)),
                  currency_fifo_amount, currency_fifo, realized_cur],
                 ['', '', '', '', '', *[positions[e] for e in positions], '', '', '', '',
                  *[eod_price[e] for e in eod_price], acc_fifo_amount, imp_sum, not_imp, not_imp_day]])
            df = pd.concat([df, new_df])
        else:
            new_df = pd.DataFrame(
                [[date, stock_name, stock_amount, stock_price_rub, '', '', *list(' ' * len(positions)),
                  stock_fifo_amount, stock_fifo, realized_stock, ],
                 [date, currency_name, currency_amount, currency_price_rub, '', '', *list(' ' * len(positions)),
                  currency_fifo_amount, currency_fifo, realized_cur],
                 ['', '', '', '', '', *[positions[e] for e in positions]]])
            df = pd.concat([df, new_df])
    cols_df = pd.DataFrame(
        [['', '', '', '', '', 'позиции', '', '', '', *list(' ' * len(positions)), 'цена на конец дня'],
         ['', '', 'количество', 'цена руб', '', *positions, '', 'кол-во для расчёта финреза', 'цена ФИФО',
          'реал', *eod_price, 'накопл финрез', 'реал накопл', 'нереал накопл', 'нереализ дневной']])
    df = pd.concat([cols_df, df])

    print("%s seconds" % round(time.time() - start_time, 2))
    start_time = time.time()
    print('writing to file...')

    df.to_excel('out.xlsx', sheet_name='out', index=False, header=False)

    print("%s seconds" % round(time.time() - start_time, 2))
    print("%s seconds total" % round(time.time() - total_time, 2))


def get_realized(fifo_amount, fifo, price, aci):
    if fifo_amount:
        return fifo_amount * (fifo - price - aci)


def get_fifo_amount(name, amount, positions, repo):
    if name != 'РУБ' and abs(positions[name] + amount) < abs(positions[name]) and not repo:
        if abs(positions[name]) >= abs(amount):
            fifo_amount = amount
        else:
            fifo_amount = -positions[name]
    else:
        fifo_amount = None
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


if __name__ == '__main__':
    main()
