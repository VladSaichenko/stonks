import datetime

import colorama as cm
import pandas as pd
from openpyxl import load_workbook
from tqdm import tqdm

pd.options.mode.chained_assignment = None


def get_custom_intervals():
    def is_valid_interval(frm, to):
        try:
            frm = tuple(map(int, frm.split('.')))
            to = tuple(map(int, to.split('.')))
            frm = datetime.date(frm[2], frm[1], frm[0])
            to = datetime.date(to[2], to[1], to[0])

            if frm <= to:
                return True
            print(cm.Fore.RED + 'Вторая дата раньше первой.')
            return False

        except ValueError:
            print(cm.Fore.RED + 'Интервал записан не корректно.')
            return False

    def get_interval(index):
        print(cm.Fore.CYAN + f'Интервал {index}/{n}')
        frm = input('От: ')
        to = input('До: ')
        print()
        return frm, to

    n = int(input(cm.Fore.CYAN + 'Введите количество интервалов: '))

    intervals = []
    for index in range(1, n + 1):
        correct = False
        while not correct:
            frm, to = get_interval(index)
            correct = is_valid_interval(frm, to)

        intervals.append((frm, to))

    return intervals


def get_intervals():
    print('1) Max;')
    print('2) Min;')
    opt = int(input('От уровня: '))
    print()

    if opt == 1:
        find = 'max'
    elif opt == 2:
        find = 'min'

    print('1) Больше или равна;')
    print('2) Меньше или равна;')
    opt = int(input('Текущая цена по отношению к уровню max/ min: '))
    print()

    if opt == 1:
        cond = 'more'
    else:
        cond = 'less'

    print(f"Текущая цена {'больше или равна' if cond == 1 else 'меньше или равна'} к {find}.")
    perc = int(input('Укажите значение в процентах: '))
    print()

    print('Укажите временные промежутки, за которые будут найдены уровни максимумов минимумов цены.')
    print(cm.Fore.LIGHTGREEN_EX + '1) За 5 дней;')
    print(cm.Fore.LIGHTYELLOW_EX + '2) За месяц;')
    print(cm.Fore.LIGHTBLUE_EX + '3) За 3 месяца;')
    print(cm.Fore.GREEN + '4) За 6 месяцев;')
    print(cm.Fore.LIGHTCYAN_EX + '5) За год;')
    print(cm.Fore.LIGHTMAGENTA_EX + '6) За 3 года;')
    print(cm.Fore.LIGHTRED_EX + '7) За 5 лет;')
    print(cm.Fore.CYAN + '8) За 10 лет;')
    print(cm.Fore.RESET + '0) Напишите в конце если хотите добавить ещё и произвольные интервалы;')

    values = set(map(int, input('Укажите промежутки: ').split()))

    intervals = []
    for value in values:
        if value == 1:
            intervals.append(('16.10.2020', '22.10.2020'))
        elif value == 2:
            intervals.append(('22.09.2020', '22.10.2020'))
        elif value == 3:
            intervals.append(('22.07.2020', '22.10.2020'))
        elif value == 4:
            intervals.append(('22.04.2020', '22.10.2020'))
        elif value == 5:
            intervals.append(('22.10.2019', '22.10.2020'))
        elif value == 6:
            intervals.append(('20.10.2017', '22.10.2020'))
        elif value == 7:
            intervals.append(('22.10.2015', '22.10.2020'))
        elif value == 8:
            intervals.append(('22.10.2010', '22.10.2020'))
        elif value == 0:
            intervals += get_custom_intervals()

    return intervals, find, cond, perc


def get_tickers():
    config = list(load_workbook('config.xlsx').worksheets[0].values)[1:]
    tickers = list(filter(lambda t: t, [row[0] for row in config]))
    return tickers


def analyse(intervals, tickers, find, cond, perc):
    def format_date(s):
        tpl = tuple(map(int, s.split('.')))
        return datetime.date(tpl[2], tpl[1], tpl[0])

    todays_global_df = pd.read_excel('todays.xlsx')

    columns = ['Тикер', 'Компания', 'Биржа', 'Max', 'Min', f'Цена ({todays_global_df.iat[0, 3]})', f'Соотношение (от {find}),%', 'От', 'До']
    result_df = all_ticks_df = pd.DataFrame([], columns=columns)

    for i in tqdm(range(0, 27)):
        df = pd.read_csv(f'stocks/stonks_{i}.csv')

        if tickers:
            intrxns = tickers
        else:
            intrxns = tuple(df.iloc[:, 0].unique())

        for ticker in intrxns:
            todays_df = todays_global_df[todays_global_df.iloc[:, 1] == ticker]
            if todays_df.empty:
                continue

            tick_df = df[df.iloc[:, 0] == ticker]
            tick_df[tick_df.columns[1]] = tuple(map(format_date, tick_df[tick_df.columns[1]]))

            try:
                curr_price = float(todays_df.iat[0, 4])
            except IndexError:
                continue

            if tick_df.empty or (not curr_price):
                continue

            for interval in intervals:
                frm, to = format_date(interval[0]), format_date(interval[1])
                filt_df = tick_df[(frm <= tick_df.iloc[:, 1]) & (tick_df.iloc[:, 1] <= to)]
                if filt_df.empty:
                    continue

                company, market = todays_df.iat[0, 0], todays_df.iat[0, 2]

                max_price, low_price = filt_df.iloc[:, 3].max(), filt_df.iloc[:, 4].min()

                level_price = max_price if find == 'max' else low_price

                ratio = round(float((curr_price - level_price) / (level_price / 100)), 3)
                row = pd.DataFrame([[ticker, company, market, float(max_price), float(low_price), curr_price,
                       round(ratio, 2), interval[0], interval[1]]], columns=columns)
                if (cond == 'more' and perc <= ratio) or (cond == 'less' and perc >= ratio):
                    result_df = result_df.append(row)
                all_ticks_df = all_ticks_df.append(row)

    return result_df, all_ticks_df


if __name__ == '__main__':
    intervals, find, cond, perc = get_intervals()
    tickers = get_tickers()
    del get_tickers, get_intervals, get_custom_intervals, cm
    res_df, whole_df = analyse(intervals, tickers, find, cond, perc)

    writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')

    whole_df.to_excel(writer, sheet_name='Все', index=False)
    res_df.to_excel(writer, sheet_name='Избранное', index=False)

    for ticker in res_df.iloc[:, 0].unique():
        tick_df = res_df[res_df.iloc[:, 0] == ticker]
        tick_df.to_excel(writer, sheet_name=ticker, index=False)

    writer.save()
