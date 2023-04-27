import os
import xlsxwriter 
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import requests as req
from fake_useragent import UserAgent
from bs4 import BeautifulSoup as BS
import re

def find_all_file(dir):
        list_dir = os.listdir(f'C:\\Users\\koooo\\Desktop\\{dir}')
        return list_dir


def find_info(files, dir):
    all_telephone = []
    for file in files: 
        file_stat = os.stat(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}')
        if ((file_stat.st_size == 5058) or (file_stat.st_size == 5057) or (file_stat.st_size == 5059) or (file_stat.st_size == 5056) or (file_stat.st_size == 5055)):
            os.remove(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}')
            continue


        data_name = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='A')
        df_name = pd.DataFrame(data_name)
        any_name = df_name.values.tolist()
        new_list_name = np.array(any_name).flatten()

        data_bin = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='B')
        df_bin = pd.DataFrame(data_bin)
        any_bin = df_bin.values.tolist()
        new_list_bin = np.array(any_bin).flatten()
        
        data_address = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='C')
        df_address = pd.DataFrame(data_address)
        any_address = df_address.values.tolist()
        new_list_address = np.array(any_address).flatten()

        data_reg = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='D')
        df_reg = pd.DataFrame(data_reg)
        any_reg = df_reg.values.tolist()
        new_list_reg = np.array(any_reg).flatten()

        data_ruc = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='F')
        df_ruc = pd.DataFrame(data_ruc)
        any_ruc = df_ruc.values.tolist()
        new_list_ruc = np.array(any_ruc).flatten()

        data_otr = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='G')
        df_otr = pd.DataFrame(data_otr)
        any_otr = df_otr.values.tolist()
        new_list_otr = np.array(any_otr).flatten()

        data_money_2021 = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='H')
        df_2021 = pd.DataFrame(data_money_2021)
        money_2021 = df_2021.values.tolist()
        new_list_money_2021 = np.array(money_2021).flatten()

        data_money_2020 = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='I')
        df_money_2020 = pd.DataFrame(data_money_2020)
        money_2020 = df_money_2020.values.tolist()
        new_list_money_2020 = np.array(money_2020).flatten()

        data_money_2022 = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='J')
        df_money_2022 = pd.DataFrame(data_money_2022)
        money_2022 = df_money_2022.values.tolist()
        new_list_money_2022 = np.array(money_2022).flatten()

        data_nalog = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='L')
        df_nalog = pd.DataFrame(data_nalog)
        nalog = df_nalog.values.tolist()
        new_list_nalog = np.array(nalog).flatten()

        data_risk = pd.read_excel(f'C:\\Users\\koooo\\Desktop\\{dir}\\{file}', usecols='M')
        df_risk = pd.DataFrame(data_risk)
        risk = df_risk.values.tolist()
        new_list_risk = np.array(risk).flatten()

        for item in range((len(new_list_bin))):
            object_company = {
                'Полное наименование': '',
                'БИН': '',
                'Адрес': '',
                'Дата регистрации': '',
                'Руководители': '',
                'Основной вид деятельности ОКЭД': '',
                '2020': 0,
                '2021': 0,
                '2022': 0,
                'Сумма налоговых отчислений': '',
                'Риски': ''  
            }

            if ((new_list_money_2020[item] > 500000000 or new_list_money_2022[item] > 500000000 or new_list_money_2021[item] > 500000000) and new_list_nalog[item] > 15000000):
                object_company['Полное наименование'] = new_list_name[item]
                item_bin = str(new_list_bin[item])
                item_bin = re.sub("[^0-9]", "", item_bin)
                if item_bin.find('.') == 1:
                    item_bin = item_bin[:-2]
                # print(len(item_bin))
                if len(item_bin) == 9:
                    item_bin = '000' + item_bin
                elif len(item_bin) == 10:
                    item_bin = '00' + item_bin
                elif len(item_bin) == 11:
                    item_bin = '0' + item_bin
                object_company['БИН'] = item_bin
                object_company['Адрес'] = new_list_address[item]
                object_company['Дата регистрации'] = new_list_reg[item]
                object_company['Руководители'] = new_list_ruc[item]
                object_company['Основной вид деятельности ОКЭД'] = new_list_otr[item]
                object_company['2020'] = new_list_money_2020[item]
                object_company['2021'] = new_list_money_2021[item]
                object_company['2022'] = new_list_money_2022[item]
                object_company['Сумма налоговых отчислений'] = new_list_nalog[item]
                object_company['Риски'] = new_list_risk[item]

                # print(object_company)
                all_telephone.append(object_company)
    # print(all_telephone)
    return all_telephone
# find_info(find_all_file('Алматы'), 'Алматы')


def create_file(file_name):
    try:
        book = xlsxwriter.Workbook(f'C:\\Users\\koooo\\Desktop\\сортировка большие компаний {file_name}.xlsx')
        page = book.add_worksheet('')
        row = 1
        column = 0
        page.set_column('A:A', 50)
        page.set_column('B:B', 50)
        page.set_column('C:C', 50)
        page.set_column('D:D', 50)
        page.set_column('F:F', 50)
        page.set_column('G:G', 50)
        page.set_column('H:H', 50)
        page.set_column('I:I', 50)
        page.set_column('J:J', 50)
        page.set_column('K:K', 50)
        page.set_column('L:L', 50)
        page.set_column('M:M', 50)

        for item in find_info(find_all_file(file_name), file_name):
            print(item)
            page.write(row, column, item['Полное наименование'])
            page.write(row, column+1, item['БИН'])
            page.write(row, column+2, item['Адрес'])
            page.write(row, column+3, item['Дата регистрации'])
            page.write(row, column+4, item['Руководители'])
            page.write(row, column+5, item['Основной вид деятельности ОКЭД'])
            page.write(row, column+6, item['2020'])
            page.write(row, column+7, item['2021'])
            page.write(row, column+8, item['2022'])
            page.write(row, column+9, item['Риски'])
            row += 1   
    except Exception as ex:
        print(ex)
    finally:
        book.close()

def all_function(dir):
    list_dir = os.listdir(f'C:\\Users\\koooo\\Desktop\\{dir}')
    for name_file in list_dir:
        create_file(name_file)

all_function('проверка')
    