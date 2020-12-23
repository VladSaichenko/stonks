import datetime

import colorama as cm
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import matplotlib.ticker as plticker
import pandas as pd
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

    return intervals


def get_tickers():
    config = list(load_workbook('graph.xlsx').worksheets[0].values)[1:]
    tickers = list(filter(lambda t: t, [row[0] for row in config]))
    return tickers


def analyse(intervals, tickers):
    def format_date(s):
        tpl = tuple(map(int, s.split('.')))
        return datetime.date(tpl[2], tpl[1], tpl[0])

    todays_global_df = pd.read_excel('todays.xlsx')

    for i in tqdm(range(0, 27)):
        df = pd.read_csv(f'stocks/stonks_{i}.csv')

        if tickers:
            intrxns = tickers
        else:
            intrxns = df.iloc[:, 0].unique()

        for ticker in intrxns:
            todays_df = todays_global_df[todays_global_df.iloc[:, 1] == ticker]
            if todays_df.empty:
                continue

            tick_df = df[df.iloc[:, 0] == ticker]
            tick_df.iloc[:, 1] = tick_df.iloc[:, 1].apply(format_date)

            try:
                curr_price = float(todays_df.iat[0, 4])
            except IndexError:
                continue

            if tick_df.empty or (not curr_price):
                continue

            # MATPLOTLIB PLOT
            plt.style.use('seaborn')
            fig, ax = plt.subplots()
            fig.subplots_adjust(bottom=0.3)
            plt.xticks(rotation=90)
            plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            xloc = plticker.MaxNLocator(nbins=31)
            ax.xaxis.set_major_locator(xloc)
            yloc = plticker.MaxNLocator(nbins=11)
            ax.yaxis.set_major_locator(yloc)
            dates = tick_df.iloc[:, 1]
            plt.plot(dates, tick_df.iloc[:, 5], linestyle='dashed', linewidth=0.25)
            plt.plot(dates, tick_df.iloc[:, 3], linestyle='solid', linewidth=0.25)
            plt.plot(dates, tick_df.iloc[:, 4], linestyle='solid', linewidth=0.25)
            tick_df = tick_df.append(pd.DataFrame([[ticker, format_date(todays_df.iat[0, 3]), None, curr_price, curr_price, curr_price, None, None]], columns=tick_df.columns))
            ax.plot(tick_df.iat[-1, 1], tick_df.iat[-1, 5], 'ro', markersize=4.3)
            plt.title(f'{ticker} history')

            for interval in intervals:
                frm, to = format_date(interval[0]), format_date(interval[1])
                filt_df = tick_df[(frm <= tick_df.iloc[:, 1]) & (tick_df.iloc[:, 1] <= to)]
                if filt_df.empty:
                    continue

                max_price = filt_df.iloc[:, 3].max()
                low_price = filt_df.iloc[:, 4].min()

                plt.hlines(max_price, xmin=frm, xmax=to, color='blue')
                plt.hlines(low_price, xmin=frm, xmax=to, color='green')

            plt.tight_layout()
            plt.savefig(f'image-graphics/{ticker}.png', dpi=380)
            plt.close('all')


if __name__ == '__main__':
    intervals = get_intervals()
    tickers = get_tickers()
    analyse(intervals, tickers)
