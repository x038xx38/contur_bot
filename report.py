# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from seleniumwire import webdriver
from selenium.common.exceptions import NoSuchElementException

import time
from datetime import date, datetime, timedelta

import requests
from lxml import html

import logging
import pprint

import function

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
from email.utils import formataddr

import sheetClass

import os


def send_mail(order_number, name, mail_content, depatment, f_name):
    # Логин и пароль для почты
    sender_address = 'svetdrus38@gmail.com'
    sender_name = Header('Иван Смотрящий', 'utf-8').encode()
    sender_pass = 'su7caCi1'
    receiver_address = ['x038xx38@gmail.com', 'r.chaplyankov@svetdrus.ru']
    # receiver_address = ['svetdrus@mail.ru', 'x038xx38@gmail.com']

    # Setup the MIME
    message = MIMEMultipart()
    message['From'] = formataddr((sender_name, sender_address))
    message['To'] = Header(', '.join(receiver_address), 'utf-8').encode()
    message['Subject'] = Header(depatment + ' :: ' + order_number + ' :: ' + name, 'utf-8')

    # Тело и вложение для письма
    message.attach(MIMEText(mail_content, 'html'))
    attach_file_name = 'orders/' + f_name
    # attach_file_name = 'C:/orders/' + order_number[7:] + '.pdf' # windows

    attach_file = open(attach_file_name, 'rb')  # открываем файл в бинарном режиме
    payload = MIMEBase('application', 'octate-stream', )
    payload.set_payload(attach_file.read())
    encoders.encode_base64(payload)  # encode the attachment
    attach_file.close()
    # add payload header with filename
    f_name = Header(f_name, 'utf-8').encode()
    payload.add_header('Content-Disposition', 'attachment; filename=%s' % f_name)
    message.attach(payload)

    # Создание SMTP сессии для отправки письма
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.starttls()
    session.login(sender_address, sender_pass)
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    msg = 'Письмо отправлено'
    return msg


def get_list_orders(stage, cookie):
    if stage == 'new':
        referer = 'https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/NewApplications'
        url = 'https://edi.kontur.ru/internal-api/supplier-web/filters/973fa598-96bf-495c-ac98-1ac3a4cd7de2/new-applications/find'
        # dt = datetime.utcnow().strftime('%Y-%m-%dT00:00:00.000') + 'Z'
        dt = (datetime.utcnow() - timedelta(days=1)).strftime('%Y-%m-%dT00:00:00.000') + 'Z'
        data = '{"filter":{"deliveryDateRange":{"lowerBound":null,"upperBound":null},"ordersNumber":null,"buyerPartyId":null,"deliveryPoint":null,"onlyUnread":false,"messageLastActivityDateRange":{"lowerBound":"' + dt + '","upperBound":null}},"sort":{"field":"ReceivedDateTime","sortOrder":"Descending"},"from":0,"size":100}'

    elif stage == 'process':
        referer = 'https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/ProcessedApplications'
        url = 'https://edi.kontur.ru/internal-api/supplier-web/filters/973fa598-96bf-495c-ac98-1ac3a4cd7de2/processed-applications/find'
        dt = (datetime.utcnow() - timedelta(days=3)).strftime('%Y-%m-%dT00:00:00.000') + 'Z'
        data = '{"filter":{"deliveryDateRange":{"lowerBound":null,"upperBound":null},"ordersNumber":null,"buyerPartyId":null,"deliveryPoint":null,"messageLastActivityDateRange":{"lowerBound":"' + dt + '","upperBound":null}},"sort":{"field":"ProcessedDateTime","sortOrder":"Descending"},"from":0,"size":40}'

    headers = {'referer': referer}
    cookies = {'cookie': cookie}
    response = requests.post(url, headers=headers, cookies=cookies, data=data)
    return response.json()


def get_list_process_orders(cookie, start, end):
    referer = 'https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/ProcessedApplications'
    url = 'https://edi.kontur.ru/internal-api/supplier-web/filters/973fa598-96bf-495c-ac98-1ac3a4cd7de2/processed-applications/find'
    # dt = (datetime.utcnow()).strftime('%Y-%m-%dT00:00:00.000') + 'Z'
    dt = '2020-01-31T16:00:00.000Z'
    data = '{"filter":{"deliveryDateRange":{"lowerBound":"'+start+'","upperBound":"'+end+'"},"ordersNumber":null,"buyerPartyId":null,"deliveryPoint":null,"messageLastActivityDateRange":{"lowerBound":"' + start + '","upperBound":null}},"sort":{"field":"ProcessedDateTime","sortOrder":"Descending"},"from":0,"size":84}'

    headers = {'referer': referer}
    cookies = {'cookie': cookie}
    response = requests.post(url, headers=headers, cookies=cookies, data=data)
    return response.json()


def get_gooditem_list(html):
    good_items = html.xpath('//div[@class="arrayField__itemData"]')
    data_order = []
    gtin = name = orders_quantity = orders_unit = price = ''
    for i in range(0, len(good_items)):
        gtin = html.xpath('*//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_GTIN"]/text()')[0]
        name = html.xpath('//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_Name"]/text()')[0]
        price = html.xpath('//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_Price"]/text()')[0]
        orders_quantity = html.xpath('//span[@id="GoodItemList_GoodItems_' + str(
            i) + '_Value_ViewModel_OrdersQuantity_CurrentValue"]/text()')[0]
        orders_unit = \
            html.xpath('//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_OrdersUnit"]/text()')[0]
        orders_quantity = orders_quantity.replace('.', ',')
        price = price.replace('.', ',')
        data_order.append(['', '', '', '', gtin, name, orders_quantity, orders_unit, price])
    return data_order


def check_stage(html):
    '''
    stage = 1 счет и документы
    stage = 2 приемка
    stage = 3 отгрузка
    '''

    stage = 2
    n_process = html.xpath('//*[contains(@id,"NProcess")]')
    n = len(n_process) - stage
    tab = html.xpath('//span[@id="NProcess_' + str(n) + '"]/@class')[0]
    return tab


def clear_sheet(credentials, sheet_id, range):
    ss = sheetClass.Spreadsheet(credentials, debugMode=False)
    ss.set_spreadsheetById(sheet_id)
    ss.clear(range)


def data_to_sheets(credentials, sheet_id, title, pos, data):
    ss = sheetClass.Spreadsheet(credentials, debugMode=False)
    ss.set_spreadsheetById(sheet_id)

    ss.append(title, data)
    ss.batch_update_values()


def main():
    logname = 'log/bot' + date.today().strftime('%d%m%Y') + '.log'
    logging.basicConfig(format='[%(asctime)s] %(filename)s:%(lineno)d %(levelname)s - %(message)s', filemode='a',
                        level=logging.INFO, filename=logname, datefmt='%d.%m.%Y %H:%M:%S')
    logging.info('Run https://edi.kontur.ru')
    start_time = time.time()
    # ------------------------------------------------------------------------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-proxy-server')
    chrome_options.add_argument('--proxy-server="direct://"')
    chrome_options.add_argument('--proxy-bypass-list=*')
    chrome_options.add_argument('user-data-sir=selenium')
    """
    Настройка для windows
    """
    # prefs = {
    #     "download.default_directory": r"C:\orders",
    #     "download.prompt_for_download": False,
    #     "download.directory_upgrade": True,
    #     "safebrowsing.enabled": True
    # }

    '''
    Настройка загрузки для MacOS
    '''
    prefs = {'download.default_directory': 'orders'}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(chrome_options=chrome_options)

    url = 'https://edi.kontur.ru/?p=1210'
    driver.get(url)
    driver.find_element_by_xpath('//a[@data-tid="tab_login"]').click()
    login = driver.find_element_by_xpath('//span/input[@data-tid="i-login"]')
    password = driver.find_element_by_xpath('//span/input[@data-tid="i-password"]')
    login.send_keys('svetdrus@mail.ru')
    password.send_keys('Brokolli2020')
    logging.info(driver.current_url)

    driver.find_element_by_xpath('//button[@type="submit"]').click()
    time.sleep(1)  # делаем небольшую паузу
    driver.refresh()
    logging.info(driver.current_url)

    response = ''
    response = driver.get_cookies()
    logging.info('Полученные куки - %s' % response)

    cookie = ''
    for i in range(0, len(response)):
        cookie = response[i]['name'] + '=' + response[i]['value'] + '; ' + cookie
    logging.info('Формат куки для запроса - %s' % cookie)
    logging.info('Вход в кабинет выполнен за  %s' % (time.time() - start_time))

    # адреса грузополучателей
    slata_gln = ['4610015765274', '4610015760026', '4610015770629', '4610015761085', '4610015762853']
    lenta_gln = ['4606068901264', '4606068902360', '4606068901103']

    neworder = []
    slata_ord = []
    lenta_ord = []

    credentials = 'ivanbot-266714-324fc59241a9.json'
    now_month = '1aoXkS4LI1-utP700z6CXYHAggGJm-biAYOxZluDWgH8' # '1JZcgnZx1YIQ-1AqMtzs-IBGcb9vnmRFG54ko--y_u5U'
    slata_sheet = '1tsEMZZil6ksZTCjPcR16PBdN8GKqeH8Sxyzt_4MBRxg'  # таблица ответов на форму Cлата
    # sheet_id_lenta = '1jZGUMzOtiUIjjNIGkv95MONoGQ2og3M9TyGQx0Mri_k' # таблица ответов на форму Лента
    # актуальные заказы, которые не прошли Приемку
    # листы Слата, Лента, Командор, Островки
    actual_orders = '1vf7T8c5WyhlgpyoMGmsBYMXJ7MyD652y7mM5hbTfTlE'  # актуальные заказы, которые не прошли Приемку

    logging.info('Проверка заказов - %s' % (time.time() - start_time))
    # order_number = tree.xpath('//span[contains(@class,"n-title-main")]/text()')[0]
    # order_number = ' '.join(order_number.split())
    # order_number = order_number[7:]
    # dt = tree.xpath('//span[@id="DeliveryDateTime_Date"]/text()')[1]
    # tm = tree.xpath('//span[@id="DeliveryDateTime_Time"]/text()')[0]
    # # delivery_date = tree.xpath('//span[@id="DeliveryDateTime_Date"]/text()')[0]
    # delivery_date = dt + ' ' + tm
    # name = tree.xpath('//span[@id="DeliveryParty_ViewModel_Name"]/text()')[0]
    # address = tree.xpath('//span[@id="DeliveryParty_ViewModel_Address"]/text()')[0]
    # gln = tree.xpath('//span[@id="DeliveryParty_ViewModel_Gln"]/text()')[0]
    # logging.info('%s :: %s :: %s :: %s :: %s' % (order_number, delivery_date, name, address, gln))
    # data_sheet = []
    # data_order = get_gooditem_list(tree)
    # for line in data_order:
    #     line[0] = order_number
    #     line[1] = delivery_date
    #     line[2] = name
    #     line[3] = address
    #     data_sheet.append(line)
    #     logging.info(line)
    # if gln in ['4610015761085', '4610015762853']:
    #     data_to_sheets(credentials, actual_orders, 'Слата', 'A', data_sheet)
    # if row[0] == '4606068999995':   # Лента
    #     data_to_sheets(credentials, actual_orders, 'Лента', 'A', data_sheet)
    # if row[0] == '4670014789992':   # Командор
    #     data_to_sheets(credentials, actual_orders, 'Командор', 'A', data_sheet)
    # if gln in ['4610015760026', '4610015770629', '4610015765274']:
    #     data_to_sheets(credentials, actual_orders, 'Островки', 'A', data_sheet)

    # ------------------------------------------------------------------------------------------------------------------
    start_time = time.time()
    start_dt = '2020-04-01T00:00:00.000Z'
    end_dt = '2020-04-30T23:59:59.000Z'
    logging.info('Делаем выгрузку всех заказов по указанному интервалу')
    datahttp = get_list_process_orders(cookie, start_dt, end_dt)

    for i in range(0, datahttp['totalCount']):
        order_id = datahttp['webFilters'][i]['info']['orderId']
        order_node_id = datahttp['webFilters'][i]['info']['orderNodeId']
        order_number = datahttp['webFilters'][i]['info']['ordersNumber']
        deliveryDate = datahttp['webFilters'][i]['info']['deliveryDate']
        partyName = datahttp['webFilters'][i]['buyerParty']['partyName']

        logging.info('%s :: %s :: %s' % (order_id, order_number, deliveryDate ))

        url = 'https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/Order/ToWebObject?orderId=' \
              + order_id + '&orderNodeId=' + order_node_id
        cookies = {'cookie': cookie}
        response = requests.get(url, cookies=cookies, allow_redirects=False)
        location = response.headers['Location']
        if location.find('Ordrsp') != -1:
            url = 'https://edi.kontur.ru'
            response = requests.get(url + location, cookies=cookies)
            tree = html.fromstring(response.content)
            link = tree.xpath('//a[@id="NProcess_0"]/@href')[0]
            response = requests.get(url + link, cookies=cookies, allow_redirects=False)
            location = response.headers['Location']

        logging.info(location)
        index = location.find('=')
        order_id = location[index+1:]
        logging.info(order_id)

        url = 'https://edi.kontur.ru'
        driver.get(url + location)
        driver.refresh()

        data_order = []
        logging.info('Сначала смотрим План ...')
        tree = html.fromstring(driver.page_source)
        delivery_date = tree.xpath('//span[@id="DeliveryDateTime_Date"]/text()')[0]
        viewmodel_name = tree.xpath('//span[@id="DeliveryParty_ViewModel_Name"]/text()')[0]
        try:
            items = tree.xpath('//div[@class="arrayField__itemData"]')
            logging.info('Количество товаров в заказе (план) - %s' % len(items))

            for j in range(0, len(items)):
                gtin = tree.xpath('*//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_GTIN"]/text()')[0]
                name = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_Name"]/text()')[0]
                price = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_Price"]/text()')[0]
                price = price.replace(' ', '')
                price = price.replace(',', '.')
                price_total_with_vat = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_PriceTotalWithVat_CurrentValue"]/text()')[0]
                price_total_with_vat = price_total_with_vat.replace(' ', '')
                price_total_with_vat = price_total_with_vat.replace(',', '.')

                price_total_vat = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_PriceTotalVat_CurrentValue"]/text()')[0]
                price_total_vat = price_total_vat.replace(' ', '')
                price_total_vat = price_total_vat.replace(',', '.')

                vat = float(price_total_with_vat) * 100/(float(price_total_with_vat)-float(price_total_vat))
                price = round(float(price)*vat/100, 1)

                orders_quantity = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_OrdersQuantity_CurrentValue"]/text()')[0]
                orders_quantity = orders_quantity.replace(' ', '')
                orders_quantity = orders_quantity.replace(',', '.')
                orders_quantity = float(orders_quantity)
                orders_unit = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_OrdersUnit"]/text()')[0]

                data_order.append([j+1, gtin, name, orders_quantity, orders_unit, price])
                # logging.info('%s : %s : %s : %s : %s : %s' % (j+1, gtin, name, orders_quantity, orders_unit, price))

        except NoSuchElementException:
            logging.error('В заказе какая-то ошибка')
        except IndexError as error:
            logging.error(error)
            driver.find_element_by_xpath('//a[@id="NProcess_0"]').click()
            driver.refresh()
            logging.info('Перешли на вкладку "Заявка"')
            tree = html.fromstring(driver.page_source)

            items = tree.xpath('//div[@class="arrayField__itemData"]')
            logging.info('Количество товаров в заказе (план) - %s' % len(items))

            for j in range(0, len(items)):
                gtin = tree.xpath('*//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_GTIN"]/text()')[0]
                name = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_Name"]/text()')[0]

                price = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_Price"]/text()')[0]
                price = price.replace(' ', '')
                price = price.replace(',', '.')
                price_total_with_vat = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_PriceTotalWithVat_CurrentValue"]/text()')[0]
                price_total_with_vat = price_total_with_vat.replace(' ', '')
                price_total_with_vat = price_total_with_vat.replace(',', '.')

                price_total_vat = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_PriceTotalVat_CurrentValue"]/text()')[0]
                price_total_vat = price_total_vat.replace(' ', '')
                price_total_vat = price_total_vat.replace(',', '.')

                vat = float(price_total_with_vat) * 100/(float(price_total_with_vat)-float(price_total_vat))
                price = round(float(price)*vat/100, 1)

                orders_quantity = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_OrdersQuantity_CurrentValue"]/text()')[0]
                orders_quantity = orders_quantity.replace(' ', '')
                orders_quantity = orders_quantity.replace(',', '.')
                orders_quantity = float(orders_quantity)
                orders_unit = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_OrdersUnit"]/text()')[0]

                data_order.append([j+1, gtin, name, orders_quantity, orders_unit, price])
                # logging.info('%s : %s : %s : %s : %s : %s' % (j+1, gtin, name, orders_quantity, orders_unit, price))

        logging.info('Теперь смотрим Факт ...')
        # делаю обертку, тк вкладка данных по Факту может быть недоступна
        try:
            # обработка вкладки Приемка
            n_process = tree.xpath('//*[contains(@id,"NProcess")]')
            n = len(n_process) - 1
            driver.find_element_by_xpath('//a[@id="NProcess_' + str(n) + '"]').click()

            tree = html.fromstring(driver.page_source)
            items = tree.xpath('//div[@class="arrayField__itemData"]')
            logging.info('Количество товаров в заказе (факт) - %s' % len(items))

            fdata_order = {}

            for j in range(0, len(items)):
                gtin = tree.xpath('*//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_GTIN"]/text()')[0]
                invoic_quantity = tree.xpath('//span[@id="GoodItemList_GoodItems_' + str(j)
                                           + '_Value_ViewModel_InvoicQuantity_CurrentValue"]/text()')[0]
                invoic_quantity = invoic_quantity.replace(' ', '')
                invoic_quantity = invoic_quantity.replace(',', '.')
                fdata_order[gtin] = invoic_quantity
            logging.info(fdata_order)

            for num in data_order:
                num.append(0.00)
                for k, v in fdata_order.items():
                    if num[1] == k:
                        v = v.replace(' ', '')
                        v = v.replace(',', '.')
                        num[6] = float(v)
                        break
                minus = (float(num[3]) - float(num[6])) * num[5]
                num.append(minus)
                num.insert(0, order_number)
                num.insert(1, viewmodel_name)
                num.insert(2, delivery_date)
                logging.info(num)
            logging.info('Номер заказа - %s' % order_number)
            logging.info('Клиент - %s' % viewmodel_name)
            logging.info('Дата доставка - %s' % delivery_date)

            data_to_sheets(credentials, now_month, 'Data', 'A', data_order)

        except NoSuchElementException:
            logging.error('Приемка по заказу %s не проходила' % order_number)

    logging.info('Формирование данных нового заказа - %s' % (time.time() - start_time))

if __name__ == '__main__':
    main()