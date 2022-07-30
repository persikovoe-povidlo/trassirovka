import pandas as pd
import random as r


def main():
    n = 6
    date = '2022-01-01'
    positions = {'RUB': 0, 'Suek': 0, 'USD': 0, 'AAPL': 0}
    eod_price = {'Suek': 0, 'USD': 0, 'AAPL': 0}
    df = pd.DataFrame([['', '', '', '', '', 'позиции', '', '', '', '', '', '', 'цена на конец дня'],
                       ['', '', 'количество', 'цена руб', '', *positions, '', 'кол--во для расчёта финреза',
                        'цена ФИФО',
                        'реал', *eod_price, 'накопл накопл', 'реал финрез', 'нереал финрез', 'нереализ дневной']])
    for i in range(n):
        m = r.randint(1,1)
        eod_price = {'Suek': r.randint(9, 11),
                     'AAPL': r.randint(20, 25),
                     'USD': r.randint(45, 55)}
        for j in range(m):
            stock_name = r.choice(['Suek', 'AAPL'])
            stock_amount = r.randint(1, 30) * 10 * r.choice([-1, 1])
            stock_price = 0
            if stock_name == 'Suek':
                stock_price = r.randint(9, 11)
            elif stock_name == 'AAPL':
                stock_price = r.randint(20, 25)

            currency_price = 0
            currency_name = r.choice(['USD', 'RUB'])
            currency_amount = r.randint(1, 30) * 10 * r.choice([-1, 1])
            if currency_name == 'USD':
                currency_price = r.randint(45, 55)
            elif currency_name == 'RUB':
                currency_price = 1

            new_part = pd.DataFrame([[date, stock_name, stock_amount, stock_price],
                                     [date, currency_name, currency_amount, currency_price], ])
            df = pd.concat([df, new_part])
            if j < m - 1:
                df = pd.concat([df, pd.DataFrame([[]])])
            else:
                df = pd.concat(
                    [df, pd.DataFrame(
                        [['', '', '', '', '', *list(' ' * len(positions)), '', '', '', '', eod_price['Suek'],
                          eod_price['USD'],
                          eod_price['AAPL']]])])
        date = str(pd.to_datetime(date) + pd.Timedelta(days=1)).split()[0]

    df.to_excel('trassirovka_generated.xlsx', sheet_name='1', index=False, header=False)


if __name__ == '__main__':
    main()
