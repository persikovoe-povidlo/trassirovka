import pandas as pd


def main():
    x = pd.read_excel('Trassirovka.xlsx', sheet_name='три инструмента', usecols=[0, 1, 2, 3, 12, 13], header=None)
    positions = {'RUB': 0, 'Suek': 0, 'USD': 0}
    queues = {'Suek': [], 'RUB': [], 'USD': []}
    eod_price = {'Suek': 0, 'USD': 0}
    imp_sum = 0
    not_imp_last_day = 0

    df = pd.DataFrame([['', '', '', '', '', 'позиции', '', '', '', '', '', '', 'цена на конец дня'],
                       ['', '', 'количество', 'цена руб', '', *positions,'', 'кол--во для расчёта финреза', 'цена ФИФО',
                        'реал', *eod_price, 'накопл финрез', 'реал финрез', 'нереал финрез', 'нереализ дневной']])

    for i in range(2, x[0].size, 3):
        stock_name = str(x[1][i])
        stock_amount = round(x[2][i], 9)
        stock_price = round(x[3][i], 9)
        currency_name = str(x[1][i + 1])
        currency_amount = round(x[2][i + 1], 9)
        currency_price = round(x[3][i + 1], 9)
        date = str(x[0][i]).split()[0]
        eod_price = {'Suek': x[12][i + 2],
                     'USD': x[13][i + 2]}

        stock_fifo = None
        currency_fifo = None

        stock_fin_res = get_fin_res(stock_name, stock_amount, positions)
        currency_fin_res = get_fin_res(currency_name, currency_amount, positions)

        positions[stock_name] += stock_amount
        positions[currency_name] += currency_amount
        queues[stock_name].append([stock_amount, stock_price])
        queues[currency_name].append([currency_amount, currency_price])

        stock_fifo = get_fifo(stock_fin_res, stock_name, stock_amount, queues)
        currency_fifo = get_fifo(currency_fin_res, currency_name, currency_amount, queues)

        if stock_fin_res:
            stock_fin_res = round(stock_fin_res, 9)
        if stock_fifo:
            stock_fifo = round(stock_fifo, 9)
        if currency_fin_res:
            currency_fin_res = round(currency_fin_res, 9)
        if currency_fifo:
            currency_fifo = round(currency_fifo, 9)

        imp_stock = get_implemented(stock_fin_res, stock_fifo, stock_price)
        imp_cur = get_implemented(currency_fin_res, currency_fifo, currency_price)
        if imp_stock:
            imp_sum += imp_stock
        if imp_cur:
            imp_sum += imp_cur

        acc_fin_res = round(
            positions['RUB'] + positions['Suek'] * eod_price['Suek'] + positions['USD'] * eod_price['USD'], 9)
        not_imp = acc_fin_res - imp_sum

        not_imp_day = not_imp - not_imp_last_day
        not_imp_last_day = not_imp

        new_df = pd.DataFrame(
            [[date, stock_name, stock_amount, stock_price, '', '', '', '', '', stock_fin_res, stock_fifo, imp_stock, '',
              '', '', '', '', ''],
             [date, currency_name, currency_amount, currency_price, '', '', '', '', '', currency_fin_res, currency_fifo,
              imp_cur, '',
              '', '', '', '', ''],
             ['', '', '', '', '', positions['RUB'], positions['Suek'], positions['USD'], '', '', '', '',
              eod_price['Suek'], eod_price['USD'], acc_fin_res, imp_sum, not_imp, not_imp_day]])
        df = pd.concat([df, new_df])
        df.to_excel('out.xlsx', sheet_name='out', index=False, header=False)


def get_implemented(fin_res, fifo, price):
    if fin_res:
        return fin_res * (fifo - price)


def get_fin_res(name, amount, positions):
    if name != 'RUB' and abs(positions[name] + amount) < abs(positions[name]):
        if abs(positions[name]) >= abs(amount):
            fin_res = amount
        else:
            fin_res = -positions[name]
    else:
        fin_res = None
    return fin_res


def get_fifo(fin_res, name, _amount, queues):
    fifo = None
    if fin_res:
        amount = 0
        b = []
        s = 0
        if queues[name][0][0] * fin_res < 0:
            if abs(queues[name][0][0]) < abs(fin_res):
                while abs(amount) < abs(fin_res):
                    if queues[name][0][0] * fin_res < 0:
                        b.append(queues[name][0])
                        queues[name].pop(0)
                        s += b[-1][1] * abs(b[-1][0])
                        amount += b[-1][0]
            else:
                amount += queues[name][0][0]
                b.append(queues[name][0])
                s += b[-1][1] * abs(b[-1][0] - (fin_res + amount))
                queues[name].pop(-1)
            dif = _amount + amount
            if dif:
                queues[name][0][0] = dif
            else:
                queues[name].pop(0)
            fifo = s / abs(fin_res)
    return fifo


if __name__ == '__main__':
    main()
