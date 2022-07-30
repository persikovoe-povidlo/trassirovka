import pandas as pd


def main():
    positions = ['RUB', 'Suek', 'USD', 'AAPL']
    n = 9 + len(positions)
    eod_cols = [i + n for i in range(len(positions) - 1)]
    x = pd.concat(
        [pd.read_excel('trassirovka_generated_.xlsx', sheet_name='1', usecols=[0, 1, 2, 3, *eod_cols], header=None),
         pd.DataFrame([[]])], ignore_index=True)
    positions = {e: 0 for e in positions}
    queues = {e: [] for e in positions}
    eod_price = dict(positions)
    eod_price.pop('RUB')
    imp_sum = 0
    not_imp_last_day = 0

    df = pd.DataFrame([['', '', '', '', '', 'позиции', '', '', '', '', '', '', 'цена на конец дня'],
                       ['', '', 'количество', 'цена руб', '', *positions, '', 'кол--во для расчёта финреза',
                        'цена ФИФО',
                        'реал', *eod_price, 'накопл накопл', 'реал финрез', 'нереал финрез', 'нереализ дневной']])

    for i in range(2, x[0].size - 1, 3):

        stock_name = str(x[1][i])
        stock_amount = round(x[2][i], 9)
        stock_price = round(x[3][i], 9)
        currency_name = str(x[1][i + 1])
        currency_amount = round(x[2][i + 1], 9)
        currency_price = round(x[3][i + 1], 9)
        date = str(x[0][i]).split()[0]
        j = n
        for e in eod_price:
            eod_price[e] = x[j][i + 2]
            j += 1

        next_date = str(x[0][i + 3]).split()[0]

        stock_fin_res = get_fin_res(stock_name, stock_amount, positions)
        currency_fin_res = get_fin_res(currency_name, currency_amount, positions)

        positions[stock_name] += stock_amount
        positions[currency_name] += currency_amount

        stock_fifo = get_fifo(stock_fin_res, stock_name, stock_amount, stock_price, queues)
        currency_fifo = get_fifo(currency_fin_res, currency_name, currency_amount, currency_price, queues)

        if stock_fin_res:
            stock_fin_res = round(stock_fin_res, 9)
        if currency_fin_res:
            currency_fin_res = round(currency_fin_res, 9)

        imp_stock = get_implemented(stock_fin_res, stock_fifo, stock_price)
        imp_cur = get_implemented(currency_fin_res, currency_fifo, currency_price)
        if imp_stock:
            imp_sum += imp_stock
        if imp_cur:
            imp_sum += imp_cur

        if date != next_date:
            acc_fin_res = positions['RUB']
            for e in eod_price:
                acc_fin_res += positions[e] * eod_price[e]
            acc_fin_res = round(acc_fin_res, 9)
            not_imp = acc_fin_res - imp_sum

            not_imp_day = not_imp - not_imp_last_day
            not_imp_last_day = not_imp
            new_df = pd.DataFrame(
                [[date, stock_name, stock_amount, stock_price, '', '', *list(' ' * len(positions)), stock_fin_res,
                  stock_fifo,
                  imp_stock],
                 [date, currency_name, currency_amount, currency_price, '', '', *list(' ' * len(positions)),
                  currency_fin_res,
                  currency_fifo, imp_cur],
                 ['', '', '', '', '', *[positions[e] for e in positions], '', '', '', '',
                  *[eod_price[e] for e in eod_price], acc_fin_res, imp_sum, not_imp, not_imp_day]])
            df = pd.concat([df, new_df])
        else:
            new_df = pd.DataFrame(
                [[date, stock_name, stock_amount, stock_price, '', '', *list(' ' * len(positions)), stock_fin_res,
                  stock_fifo,
                  imp_stock, ],
                 [date, currency_name, currency_amount, currency_price, '', '', *list(' ' * len(positions)),
                  currency_fin_res,
                  currency_fifo, imp_cur],
                 ['', '', '', '', '', *[positions[e] for e in positions]]])
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


def get_fifo(fin_res, name, amount, price, queues):
    fifo = None
    if fin_res:
        if len(queues[name]) > 1:
            fifo = 0
            for e in queues[name]:
                fifo += abs(e[0]) * e[1]
            fifo /= fin_res
        else:
            fifo = queues[name][0][1]

    if queues[name]:
        if queues[name][0][0] * amount <= 0:
            queues[name].insert(0, [amount, price])
            while len(queues[name]) > 1:
                if abs(queues[name][0][0]) < abs(queues[name][1][0]):
                    queues[name][0] = [queues[name][0][0] + queues[name][1][0], queues[name][1][1]]
                else:
                    queues[name][0] = [queues[name][0][0] + queues[name][1][0], queues[name][0][1]]
                queues[name].pop(1)
        else:
            queues[name].append([amount, price])
    else:
        queues[name].append([amount, price])
    return fifo


if __name__ == '__main__':
    main()
