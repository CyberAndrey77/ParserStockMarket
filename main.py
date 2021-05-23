# coding: utf8
import datetime
import time
import requests
from lxml import etree
from bs4 import BeautifulSoup
import re
import fileManager
import emailSender


def cleanhtml(raw_html):
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, '', raw_html)
    return cleantext

def build_list(table, rows):
    row = []
    j = 0
    while j < len(table):
        for item in table[j]:
            row.append(cleanhtml(str(item)))
        rows.append(row)
        row = []
        j += 1
    return rows

def get_html(url):
    response = requests.get(url)
    return response.text

def parse_html(html):
    soup = BeautifulSoup(html, 'lxml')
    table = soup.find_all("tr", class_ = 'tr1')
    rows = []
    rows = build_list(table, rows)

    table = soup.find_all("tr", class_='tr0')

    rows = build_list(table, rows)

    extraCharacter = '\n'
    indexDate = 0
    indexCurrency = 3

    for row in rows:
        for item in row:
            if item == extraCharacter:
                row.remove(item)

    for row in rows:
        row[indexDate] = datetime.datetime.strptime(row[indexDate], '%d.%m.%Y').date()
        row[indexCurrency] = row[indexCurrency].replace(',', '.')
        try:
            row[indexCurrency] = float(row[indexCurrency])
        except ValueError:
            row[indexCurrency] = 0

    for row in rows:
        for i in range(len(row) - 1, -1, -1):
            if i != indexDate and i != indexCurrency:
                del row[i]

    month = datetime.datetime.today().month

    for i in range(len(rows) - 1, -1, -1):
        if rows[i][indexDate].month < month:
            del rows[i]

    rows = sorted(rows, key=lambda x: x[indexDate], reverse=True)

    return rows

def main():
    url1 = 'https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=USD_RUB'
    url2 = 'https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=EUR_RUB'
    html = get_html(url1)
    rowsUSD = parse_html(html)
    print(rowsUSD)
    html = get_html(url2)
    rowsEUR = parse_html(html)
    print(rowsEUR)
    rows = fileManager.write_exel(rowsUSD, rowsEUR)
    textMessage = 'строк'

    if rows >= 10 and rows <=20:
        textMessage = str(rows) + ' ' + textMessage

    number = rows % 10
    if number == 1:
        textMessage = str(rows) + ' ' + textMessage + 'а'
    elif number == 2 or number == 3 or number == 4:
        textMessage = str(rows) + ' ' + textMessage + 'и'

    textMessage = 'В таблице ' + textMessage

    #emailSender.send_message('lapardin.andrey@mail.ru', textMessage, 'USD_EUR.xlsx')


if __name__ == '__main__':
    main()


