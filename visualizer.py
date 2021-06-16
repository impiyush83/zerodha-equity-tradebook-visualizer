import openpyxl
from collections import defaultdict
from pathlib import Path
import matplotlib.pyplot as plt


def file_input():
    xlsx_file = Path('', 'tradebook.xlsx')
    return openpyxl.load_workbook(xlsx_file).active


def sanitize(sheet):
    _clean_data = []
    for row in sheet.values:
        if row[1] not in (None, 'Client ID', 'Symbol', ''):
            _clean_data.append(row)
    return _clean_data


def aggregate_stocks(clean_data):
    _stock_aggregator = defaultdict(int)
    for trade in range(1, len(clean_data)):
        if clean_data[trade][7] == 'buy':
            _stock_aggregator[clean_data[trade][1]] -= (clean_data[trade][8]
                                                       * clean_data[trade][9])
        else:
            _stock_aggregator[clean_data[trade][1]] += (clean_data[trade][8]
                                                       * clean_data[trade][9])
    return _stock_aggregator


def draw(aggregated_stocks, colors):
    stocks = aggregated_stocks.keys()
    earnings = aggregated_stocks.values()
    plt.bar(stocks, earnings, color=colors)
    plt.title('Stocks Vs Earnings')
    plt.xlabel('Stocks')
    plt.ylabel('Earnings')
    plt.grid(True)
    plt.show()


if __name__ == '__main__':
    sheet = file_input()
    clean_data = sanitize(sheet)
    stock_aggregator = aggregate_stocks(clean_data)
    draw(stock_aggregator, colors = ['green', 'blue', 'purple', 'brown'])
    exit()
