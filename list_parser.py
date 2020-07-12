import xlsxwriter
import requests
from bs4 import BeautifulSoup as bs
import html2text

headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
vacancy = 'Программист'
base_url = 'https://music.yandex.ru/genre/русский%20рок/albums/new?page='
pages = 10
records = []


def ym_parse(base_url, headers):
    zero = 0
    workbook = xlsxwriter.Workbook('rus_Records.xlsx')
    worksheet = workbook.add_worksheet()
    # Добавим стили форматирования
    bold = workbook.add_format({'bold': 1})
    bold.set_align('center')
    center_H_V = workbook.add_format()
    center_H_V.set_align('center')
    center_H_V.set_align('vcenter')
    center_V = workbook.add_format()
    center_V.set_align('vcenter')
    cell_wrap = workbook.add_format()
    cell_wrap.set_text_wrap()

    # Настройка ширины колонок
    worksheet.set_column(0, 0, 40)  # A  https://xlsxwriter.readthedocs.io/worksheet.html#set_column
    worksheet.set_column(1, 1, 40)  # B
    worksheet.set_column(2, 2, 40)  # C
    worksheet.set_column(3, 3, 40)  # D
    worksheet.set_column(4, 4, 40)  # E
    worksheet.set_column(5, 5, 40)  # F

    worksheet.write('A1', 'Релиз', bold)
    worksheet.write('B1', 'Группа', bold)
    worksheet.write('C1', 'Год', bold)
    worksheet.write('D1', 'Ссылка на релиз', bold)
    worksheet.write('E1', 'Лейбл', bold)

    while pages > zero:
        zero = str(zero)
        session = requests.Session()
        request = session.get(base_url + zero, headers=headers)
        if request.status_code == 200:
            soup = bs(request.content, 'html.parser')
            divs = soup.find_all('div', attrs={'class': 'album'})
            for div in divs:
                title = div.find('div', attrs={'class': 'album__title'}).text
                artist = div.find('div', attrs={'class': 'album__artist'}).text
                href = div.find('a', attrs={'class': 'd-link deco-link album__caption'})['href']
                year = div.find('div', attrs={'class': 'album__year'}).text
                """text1 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'}).text
                text2 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_requirement'}).text
                content = text1 + '  ' + text2"""
                request2 = session.get('https://music.yandex.ru'+href, headers=headers)
                label = 'Лейбл не указан'
                if request2.status_code == 200:
                    soup2 = bs(request2.content, 'html.parser')
                    try:
                        divs2 = soup2.find('div', attrs={'class': 'page-album__label'}).text
                    except:
                        pass
                    label = str.lower(html2text.html2text(divs2))[5:].replace('\n', '')
                    print(label)
                    #label = divs2.find('a', attrs={'a': 'd-link'}).text
                    #label_href = divs2.find('a', attrs={'a': 'd-link'})['href']
                all_txt = [title, artist, year, href, label]
                records.append(all_txt)
            zero = int(zero)
            zero += 1

        else:
            print('error')

        # Запись в Excel файл

        row = 1
        col = 0
        for i in records:
            worksheet.write_string(row, col, i[0], center_V)
            worksheet.write_string(row, col + 1, i[1], cell_wrap)
            worksheet.write_string(row, col + 2, i[2], center_H_V)
            worksheet.write_string(row, col + 3, i[3], center_H_V)
            worksheet.write_string(row, col + 4, i[4], cell_wrap)
            # worksheet.write_url (row, col + 4, i[4], center_H_V)
            # worksheet.write_url(row, col + 5, i[5])
            row += 1

        print('OK')
    workbook.close()


ym_parse(base_url, headers)
