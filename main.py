import requests as req
from fake_useragent import UserAgent
from bs4 import BeautifulSoup as BS
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import xlsxwriter
import time
from threading import Thread
import os



def get_oblast_url() :
        
    useragent = UserAgent()
    headers = {'Uset-Agent':
                f'{useragent.random}'
            }
    
    url = 'https://statsnet.co/states/kz'
    response = req.get(url, headers=headers)
    soup = BS(response.text, 'lxml')

    ul = soup.find('ul', class_='flex flex-1 flex-col')
    a = ul.find_all('a')
    link = []
    for car in a:
        oblast = {
            'name': '',
            'link': ''
        }

        oblast['link'] = car.get('href')
        oblast['name'] = car.text
        link.append(oblast)
    
    return link

def get_company_url(oblast, start, end):
    useragent = UserAgent()
    headers = {'Uset-Agent':
                f'{useragent.random}'
            }
    link = oblast['link']
    company_link = []
    for count in range(start, end):        
        url = f'https://statsnet.co{link}?page={count}'
        response = req.get(url, headers=headers)
        soup = BS(response.text, 'lxml')
        a = soup.find_all('a', class_='text-lg sm:text-xl flex items-center gap-1 text-statsnet hover:text-orange font-stem')
        for item in a:
            company_link.append(item.get('href'))
        print(count)
    print(len(company_link))
    return company_link

def get_info_about_thecompany(oblast, start, end):
    useragent = UserAgent()
    all_item_list = []
    for link in get_company_url(oblast, start, end):
        try:
            url = f'https://statsnet.co{link}'
            options = webdriver.FirefoxOptions()
            options.set_preference("general.useragent.override", useragent.random)
            options.set_preference('dom.webdriver.enabled', False)
            options.headless = True
            service=Service(r"C:\Users\koooo\Desktop\cбор компаний со Statnet\firefoxdriver\geckodriver.exe")
            driver = webdriver.Firefox(
                service=service,
                options=options
            )

            driver.get(url)
            time.sleep(3)
            object_info = {
                'Полное наименование': '',
                'БИН': '',
                'Адрес': '',
                'Дата регистрации': '',
                'Отрасль': '',
                'Руководители': '',
                'Основной вид деятельности ОКЭД': '',
                '2020': 0,
                '2021': 0,
                '2022': 0,
                'Сумма налоговых отчислений': '',
                'Выручка с контрактов': '',
                'Риски': ''  
            }
            def search(item, name_object_item):
                if item.text.find('stat.gov.kz') == -1 and item.text.find('kgd.gov.kz') == -1:
                    object_info[name_object_item] = item.text
                elif item.text.find('stat.gov.kz'):
                    item = item.text[:-12]
                    object_info[name_object_item] = item
                elif item.text.find('kgd.gov.kz'):
                    item = item.text[:-7]
                    object_info[name_object_item] = item

            table_info = driver.find_element(By.XPATH, '/html/body/div/main/div/div[3]/table')
            trs_info = table_info.find_elements(By.TAG_NAME, 'tr')
            for tr_info in trs_info:
                tds_info = tr_info.find_elements(By.TAG_NAME, 'td')
                if (tds_info[0].text) == 'Полное наименование':
                    search(tds_info[1], 'Полное наименование')
                if (tds_info[0].text) == 'БИН':
                    search(tds_info[1], 'БИН')
                if (tds_info[0].text) == 'Адрес':
                    search(tds_info[1], 'Адрес')
                if (tds_info[0].text) == 'Дата регистрации':
                    search(tds_info[1], 'Дата регистрации')
                if (tds_info[0].text) == 'Отрасль':
                    search(tds_info[1], 'Отрасль')
                if (tds_info[0].text) == 'Руководители':
                    search(tds_info[1], 'Руководители')

            finance = driver.find_elements(By.XPATH, '/html/body/div/main/div/div[5]/div/div[2]/h4/div/div/div')
            for item in finance:
                if item.get_attribute('data-year') == '2020':
                    value = item.get_attribute('data-value')
                    # print(value)
                    if value != '' :
                        object_info['2020'] = int(value)
                if item.get_attribute('data-year') == '2021':
                    value = item.get_attribute('data-value')
                    # print(value)
                    if value != '' :
                        object_info['2021'] = int(value)
                if item.get_attribute('data-year') == '2022':
                    value = item.get_attribute('data-value')
                    # print(value)
                    if value != '' :
                        object_info['2022'] = int(value)
            Kind_of_activity = driver.find_element(By.ID, 'activity_type')
            name_activitye = Kind_of_activity.find_elements(By.CLASS_NAME, 'col-span-8')
            object_info['Основной вид деятельности ОКЭД'] = name_activitye[0].text

            nalog = driver.find_element(By.ID, 'link1').find_element(By.TAG_NAME, 'h4').text
            if (nalog == 'Не найдено') :
                object_info['Сумма налоговых отчислений'] = 0
            else :
                nalog_sort = re.sub('[^0-9]', '', nalog)
                object_info['Сумма налоговых отчислений'] = int(nalog_sort)

            contracts = driver.find_element(By.ID, 'government_contracts').find_element(By.TAG_NAME, 'span').text
            object_info['Выручка с контрактов'] = contracts

            riski = driver.find_element(By.ID, 'risks').find_elements(By.TAG_NAME, 'span')[0].text
            object_info['Риски'] = riski
            print('good')
            all_item_list.append(object_info)
        
        finally : 
            driver.close()
            driver.quit()
    return all_item_list

def writer(name_oblast, oblast, start, end):
    try:

        book = xlsxwriter.Workbook(f'C:\\Users\\koooo\\Desktop\\cбор компаний со Statnet\\{name_oblast}\\{start}-{end} сортировка {name_oblast}.xlsx')
        page = book.add_worksheet('')
        
        row = 0
        column = 0
 
        page.set_column('A:A', 50)
        page.set_column('B:B', 20)
        page.set_column('C:C', 40)
        page.set_column('D:D', 40)
        page.set_column('E:E', 20)
        page.set_column('F:F', 20)
        page.set_column('G:G', 20)
        page.set_column('H:H', 20)
        page.set_column('I:I', 20)
        page.set_column('J:J', 20)
        page.set_column('K:K', 20)
        page.set_column('L:L', 20)
        page.set_column('M:M', 20)
        for item in get_info_about_thecompany(oblast, start, end):
            page.write(row, column, item['Полное наименование'])
            page.write(row, column+1, item['БИН'])
            page.write(row, column+2, item['Адрес'])
            page.write(row, column+3, item['Дата регистрации'])
            page.write(row, column+4, item['Отрасль'])
            page.write(row, column+5, item['Руководители'])
            page.write(row, column+6, item['Основной вид деятельности ОКЭД'])
            page.write(row, column+7, item['2020'])
            page.write(row, column+8, item['2021'])
            page.write(row, column+9, item['2022'])
            page.write(row, column+10, item['Выручка с контрактов'])
            page.write(row, column+11, item['Сумма налоговых отчислений'])
            page.write(row, column+12, item['Риски'])
            row += 1
    except:
        pass
    finally:
        book.close()

def trackTime(oblast_count):
    oblast = get_oblast_url()
    start = time.time() 
    print(start)

    name = oblast[oblast_count]['name'].replace(' ', '') 
    os.mkdir(f'{name}')
    for car in range(0, 10):
        for item in range(0, 100):
            useragent = UserAgent()
            headers = {'Uset-Agent':
                        f'{useragent.random}'
                    }
            count = item+1
            if len(str(item)) == 1 and car != 0:
                item = f'0{item}'
            if len(str(count)) == 1 and car != 0:
                count = f'0{count}'
            if item == 0 and car == 0:
                start = f'{count}'
                end = f'{count}0'
                start = int(start)
                end = int(end)
                writer(name, oblast[oblast_count], start, end)
                continue
            if car == 0 and item != 0:
                start = f'{item}1'
                end = f'{count}0'
                start = int(start)
                end = int(end)
                writer(name, oblast[oblast_count], start, end)
                continue
            if len(str(count)) == 3:
                number = car + 1    
                start = f'{car}{item}1'
                end = f'{number}000'
                start = int(start)
                end = int(end)
                writer(name, oblast[oblast_count], start, end)
            else:
                start = f'{car}{item}1'
                end = f'{car}{count}0'
                start = int(start)
                end = int(end)
                writer(name, oblast[oblast_count], start, end)               
    end = time.time() - start
    end = end/60
    print(end)


def two_potok():
    for item in range(14, 20):
        if item % 2 == 1:
            continue
        count = item + 1
        t1 = Thread(target=trackTime, args=(item, ))
        t2 = Thread(target=trackTime, args=(count, ))

        t1.start()
        t2.start()

        t1.join()
        t2.join()

two_potok()