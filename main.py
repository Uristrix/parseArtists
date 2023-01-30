from bs4 import BeautifulSoup
import requests
import xlsxwriter
import os
import re
from dotenv import load_dotenv

load_dotenv()

delete_str = 'Внимание! ARTInvestment.ru ищет картины этого художника для продажи.                            '
err_name1 = 'перейти к работе'
err_name2 = 'без названия'


def parse(URL, AUCTION, NUM, LETTER):
    data = []
    page = 1
    while True:
        try:
            link = f'{URL}/{AUCTION}/?first_letter={LETTER}&page={page}&citems={NUM}'
            print(link)
            html = requests.get(link).text
            soup = BeautifulSoup(html, 'html.parser')
            elems = soup.findAll('div', {"class": "artists-list"})

            if not elems:
                break

            for el in elems:
                temp = {}
                artist_link = el.find('a', href=True)['href']
                soup_artist = BeautifulSoup(requests.get(artist_link).text, 'html.parser').find('div',
                                                                                                {'class': 'white'})

                name = soup_artist.find('h1').text.split('\n')[1]
                temp['nameRu'] = re.sub(r'^' + delete_str, '', re.sub(r"^\s+", '', name)) if name is not None else ' '

                name = soup_artist.findAll('a', {'class': 'read-all'}, href=True)
                nameEn = ' '
                for n in name:
                    if len(re.sub(r'[А-Яа-яёЁ]', '', n.text)) == len(n.text):
                        nameEn = n.text
                        break

                temp['nameEn'] = re.sub(r',', '', nameEn)

                date1 = soup_artist.find('p', {'class': 'painter-meta'})
                date2 = soup_artist.find('em', {'class': 'high'})
                temp['date'] = date1.text if date1 is not None else date2.text if date2 is not None else ' '

                spec = soup_artist.find('p', {'class': 'mat2'})
                temp['spec'] = spec.text if spec is not None else ' '

                bio = soup_artist.findAll('p', {'class': 'mat1'})
                biores = []
                for e in bio:
                    biores.append(e.text if e.text is not None else " ")
                temp['bio'] = biores

                a_work = soup_artist.find('a', {'class': 'artists-subtitle'}, href=True)

                if a_work is not None:
                    work = BeautifulSoup(requests.get(a_work['href']).text, 'html.parser').find('div',
                                                                                        {'class': 'content-data'})

                    work_soup = work.findAll('div', {'class': 'list-item'})
                    works = []
                    for w in work_soup:
                        name = w.find('h3')
                        if name is not None and \
                                str(name.text).lower() != err_name1 and \
                                str(name.text).lower() != err_name2:
                            works.append(name.text)

                    temp['works'] = works[:7]
                else:
                    temp['works'] = []

                data.append(temp)
                print(temp)
                # if soup_artist.find('a', {'id': 'bio'}) is None:
                #     continue
            page += 1
        except Exception as e:
            print(e)
            break

    return data


def create_xlsx(data, name):
    workbook = xlsxwriter.Workbook(f'{name}.xlsx')
    worksheet = workbook.add_worksheet('Data')
    style1 = workbook.add_format({
        'bold': 1,
        'border': 2,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '00AEFF',
        "font_color": "white"
    })

    style2 = workbook.add_format({
        'bold': 1,
        'border': 2,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': 'F0FFFF',
        'text_wrap': 1
    })

    # Ширина колонок
    for i, el in enumerate([30, 30, 20, 35, 60, 40]):
        worksheet.set_column(i, i, el)

    # Заголовок
    header = ['ФИО автора (rus)', 'ФИО автора (eng)', 'Период жизни', 'Специализация', 'Ключевые этапы биографии',
              'Работы автора']

    for i, el in enumerate(header):
        worksheet.write(0, i, el, style1)

    for i, el in enumerate(data):
        # имяRu
        worksheet.write(i + 1, 0, el['nameRu'], style2)

        # имяEn
        worksheet.write(i + 1, 1, el['nameEn'], style2)

        # дата
        worksheet.write(i + 1, 2, el['date'], style2)

        # специализация
        worksheet.write(i + 1, 3, el['spec'], style2)

        # Биография
        worksheet.write(i + 1, 4, '\n'.join(el['bio']), style2)

        # Работы
        worksheet.write(i + 1, 5, '\n'.join(el['works']), style2)

    workbook.close()


if __name__ == '__main__':
    # for i in range(ord('Я'), ord('А')+32):
    create_xlsx(parse(os.getenv('URL'), os.getenv('AUCTION'), int(os.getenv('NUM')), 'Ш'), 'Ш')
