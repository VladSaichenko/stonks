import csv
import datetime

import colorama as cm
from openpyxl import load_workbook
from tqdm import tqdm


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

    wb = load_workbook('todays.xlsx')

    todays_df = tuple(wb.worksheets[0].values)[1:]

    result_df = all_ticks_df = []
    for i in tqdm(range(0, 27)):
        with open(f'stocks/stonks_{i}.csv') as f:
            df = tuple(list(row) for row in csv.reader(f))[1:]

        # Formatting data
        for index, row in enumerate(df):
            df[index][1] = format_date(row[1])
            df[index][2] = float(row[2]) if row[2] else None
            df[index][3] = float(row[3]) if row[2] else None
            df[index][4] = float(row[4]) if row[2] else None
            df[index][5] = float(row[5]) if row[2] else None
            df[index][6] = float(row[6]) if row[2] else None
            df[index][7] = int(float(row[7])) if row[7] else None

        if tickers:
            intrxns = set(tickers)
        else:
            intrxns = set(row[0] for row in df)

        if None in intrxns:
            intrxns.remove(None)

        for ticker in intrxns:
            tick_df = tuple(filter(lambda r: r[0] == ticker, df))
            try:
                curr_price = tuple(filter(lambda r: r[1] == ticker, todays_df))[0][4]
            except IndexError:
                continue

            if (not tick_df) or (not curr_price):
                continue

            for interval in intervals:
                frm, to = format_date(interval[0]), format_date(interval[1])
                filt_df = tuple(filter(lambda r: (r[1] < to) and (r[1] > frm) and ((r[3] != None) and (r[4] != None) and (r[5] != None)), tick_df))
                company, _, market = tuple(filter(lambda r: r[1] == ticker, todays_df))[0][:3]

                if not filt_df:
                    continue

                max_price = max(set(row[3] for row in filt_df))
                low_price = min(set(row[4] for row in filt_df))

                if find == 'max':
                    ratio = round(float((curr_price-max_price)/(max_price/100)), 3)
                    row = [ticker, company, market, find, float(max_price), float(low_price), curr_price,
                           round(ratio, 2), interval[0], interval[1]]
                    if (cond == 'more' and perc <= ratio) or (cond == 'less' and perc >= ratio):
                        result_df.append(row)
                    all_ticks_df.append(row)

                elif find == 'min':
                    ratio = round(float((curr_price-low_price)/(low_price/100)), 3)
                    row = [ticker, company, market, find, float(max_price), float(low_price), curr_price,
                           round(ratio, 2), interval[0], interval[1]]
                    if (cond == 'more' and perc <= ratio) or (cond == 'less' and perc >= ratio):
                        result_df.append(row)
                    all_ticks_df.append(row)


        todays_price = todays_df[0][3]

    return result_df, all_ticks_df, todays_price


if __name__ == '__main__':
    intervals, find, cond, perc = get_intervals()
    tickers = get_tickers()
    del get_tickers, get_intervals, get_custom_intervals, cm
    res_df, whole_df, todays_price = analyse(intervals, tickers, find, cond, perc)

    import xlsxwriter
    workbook = xlsxwriter.Workbook('result.xlsx')
    del xlsxwriter

    def add_headers(sheet):
        sheet.write(0, 0, 'Тикер')
        sheet.write(0, 1, 'Компания')
        sheet.write(0, 2, 'Биржа')
        sheet.write(0, 3, f'Уровень {find}')
        sheet.write(0, 4, 'Max')
        sheet.write(0, 5, 'Min')
        sheet.write(0, 6, f'Цена {todays_price}')
        sheet.write(0, 7, f'Соотношение {find},%')
        sheet.write(0, 8, f'От')
        sheet.write(0, 9, f'До')

    favorites_sheet = workbook.add_worksheet('Избранное')
    add_headers(favorites_sheet)
    line = 0
    for i, row in enumerate(res_df):
        for j, val in enumerate(row):
            favorites_sheet.write(i+1, j, val)

    all_sheet = workbook.add_worksheet('Все')
    add_headers(all_sheet)
    line = 0
    for i, row in enumerate(whole_df):
        for j, val in enumerate(row):
            all_sheet.write(i+1, j, val)

    unique_ticks = set(row[0] for row in whole_df)

    for tick in unique_ticks:
        tick_sheet = workbook.add_worksheet(tick)
        add_headers(tick_sheet)
        for i, row in enumerate(tuple(filter(lambda r: r[0] == tick, whole_df))):
            for j, val in enumerate(row):
                tick_sheet.write(i+1, j, val)

    workbook.close()
