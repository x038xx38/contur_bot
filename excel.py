from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from seleniumwire import webdriver
from selenium.common.exceptions import NoSuchElementException

import sheetClass
import pprint

import time
from datetime import date
from datetime import datetime
from datetime import timedelta

import requests
from lxml import html
import json

import logging


def get_list_orders(stage, cookie, start, end):
    start = start + 'T00:00:00.000Z'
    end = end + 'T23:59:59.000Z'
    if stage == 'new':
        referer = 'https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/NewApplications'
        url = 'https://edi.kontur.ru/internal-api/supplier-web/filters/973fa598-96bf-495c-ac98-1ac3a4cd7de2/new-applications/find'
        # dt = datetime.utcnow().strftime('%Y-%m-%dT00:00:00.000') + 'Z'
        # dt = (datetime.utcnow() - timedelta(days=7)).strftime('%Y-%m-%dT00:00:00.000') + 'Z'
        data = '{"filter":{"deliveryDateRange":{"lowerBound":null,"upperBound":null},"ordersNumber":null,"buyerPartyId":null,"deliveryPoint":null,"onlyUnread":false,"messageLastActivityDateRange":{"lowerBound":"' + start + '","upperBound":"' + end + '"}},"sort":{"field":"ReceivedDateTime","sortOrder":"Descending"},"from":0,"size":100}'

    elif stage == 'process':
        referer = 'https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/ProcessedApplications'
        url = 'https://edi.kontur.ru/internal-api/supplier-web/filters/973fa598-96bf-495c-ac98-1ac3a4cd7de2/processed-applications/find'
        # dt = (datetime.utcnow() - timedelta(days=8)).strftime('%Y-%m-%dT00:00:00.000') + 'Z'
        data = '{"filter":{"deliveryDateRange":{"lowerBound":null,"upperBound":null},"ordersNumber":null,"buyerPartyId":null,"deliveryPoint":null,"messageLastActivityDateRange":{"lowerBound":"' + start + '","upperBound":"' + end + '"}},"sort":{"field":"ProcessedDateTime","sortOrder":"Descending"},"from":0,"size":40}'

    headers = {'referer': referer}
    cookies = {'cookie': cookie}
    response = requests.post(url, headers=headers, cookies=cookies, data=data)
    return response.json()


def get_gooditem_list(html):
    good_items = html.xpath('//div[@class="arrayField__itemData"]')
    data_order = []
    gtin = name = orders_quantity = orders_unit = price = ''
    logging.info('Количество товаров в заказе - %s' % len(good_items))
    for i in range(0, len(good_items)):
        gtin = html.xpath('*//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_GTIN"]/text()')[0]
        name = html.xpath('//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_Name"]/text()')[0]
        price = html.xpath('//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_Price"]/text()')[0]
        orders_quantity = html.xpath('//span[@id="GoodItemList_GoodItems_' + str(
            i) + '_Value_ViewModel_OrdersQuantity_CurrentValue"]/text()')[0]
        orders_unit = \
            html.xpath('//span[@id="GoodItemList_GoodItems_' + str(i) + '_Value_ViewModel_OrdersUnit"]/text()')[0]
        data_order.append(['', '', gtin, name, orders_quantity, orders_unit, price])

    return data_order


def deviation(plan, fact):
    '''
    Считает процентное отклоение факта от плана в заказе
    '''
    a = plan.replace(' ', '')
    b = fact.replace(' ', '')
    a = a.replace(',', '.')
    b = b.replace(',', '.')
    result = 100 - (float(b) * 100 / float(a))
    return result

def main():
    logname = 'log/excel_' + date.today().strftime('%d%m%Y') + '.log'
    logging.basicConfig(format='[%(asctime)s] %(filename)s:%(lineno)d %(levelname)s - %(message)s', filemode='a',
                        level=logging.INFO, filename=logname, datefmt='%d.%m.%Y %H:%M:%S')
    start_time = time.time()
    # ------------------------------------------------------------------------------------------------------------------
    credentials_path = 'ivanbot-266714-324fc59241a9.json'

    number = '9501304006'  # телефоный номер сотрудника, по нему происходит поиск и оценка
    start_dt = '2020-02-01'  # начальная дата выборки
    end_dt = '2020-02-28'  # конечная дата выборки

    slata_sheet = '1tsEMZZil6ksZTCjPcR16PBdN8GKqeH8Sxyzt_4MBRxg'  # таблица ответов на форму Cлата
    dataproduct_sheet = '1Kc_8fx_L5CjTp0Xc7GNi3lnQT92KnpImzhTZxo26bQM'  # стоимость работ, название, GTIN продукта
    filter_sheet = '19lsnbARjGdKr4yHlq2CPQ8vTZpE0jl4eIheLs_riA0M'  # таблица Фильтр
    ss = sheetClass.Spreadsheet(credentials_path, debugMode=False)
    ss.set_spreadsheetById(slata_sheet)
    range_sheet = 'A1:' + str(ss.rowCount)

    ss.set_spreadsheetById(filter_sheet)
    formula = '=QUERY(IMPORTRANGE("' + slata_sheet + '";"' + range_sheet + '");"select * where Col3=' + number \
              + ' and Col1 > date \'' + start_dt + '\' and Col1 < date \'' + end_dt + '\'")'
    ss.add_formula('A2:A2', formula)
    formula = '=TRANSPOSE(IMPORTRANGE("' + dataproduct_sheet + '";"Слата!B2:B200"))'
    ss.add_formula('E1:E1', formula)
    ss.batch_update_spreadsheet()

    ranges = ['filter']
    data_sheet = ss.batch_get_values(ranges)
    # pprint.pprint(data_sheet)

    workers_data = []
    size_line = len(data_sheet[0])
    for i in range(2, len(data_sheet)):
        by_order = {}
        for j in range(4, size_line):
            by_order[data_sheet[0][j]] = data_sheet[i][j]
        workers_data.append([data_sheet[i][0], data_sheet[i][2], data_sheet[i][3][7:], by_order.copy()])
    # for row in workers_data:
    #     print(row)

    done = [*workers_data[0][3]]  # ключи одной строки из заказа, предполагается что их количество везде одинаковое
    logging.info('GTIN продуктов, которые делает Цех для Слаты: %s' % done)

    ss.set_spreadsheetById(dataproduct_sheet)
    ranges = ['Слата']
    data_sheet = ss.batch_get_values(ranges)
    cost_work = {}
    for i in range(1, len(data_sheet)):
        cost_work[data_sheet[i][1]] = data_sheet[i][2][:-2].replace(',', '.')   # 2,00 р --> 2.00
    logging.info('Стоимость работы продуктов %s ' % cost_work)

    # ------------------------------------------------------------------------------------------------------------------
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-proxy-server')
    chrome_options.add_argument('--proxy-server="direct://"')
    chrome_options.add_argument('--proxy-bypass-list=*')
    chrome_options.add_argument('user-data-sir=selenium')

    '''
    Настройки загрузки для MacOS
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
    time.sleep(1)  # wait load selenium
    driver.refresh()
    logging.info(driver.current_url)

    response = ''
    response = driver.get_cookies()
    logging.info('Полученные куки - %s' % response)

    cookie = ''
    for i in range(0, len(response)):
        cookie = response[i]['name'] + '=' + response[i]['value'] + '; ' + cookie
    logging.info('Формат куки для запроса - %s' % cookie)
    # ------------------------------------------------------------------------------------------------------------------

    logging.info('Теперь прохожусь по заказам которые уже в process')
    logging.info('get_list_orders("process", cookie)')

    response = []
    response = get_list_orders('process', cookie, start_dt, end_dt)
    logging.info(response)

    buff = []
    for i in range(0, len(workers_data)):
        buff.append(workers_data[i][2])
    a_set = set(buff)
    logging.info('Номера заказов из Цеха: %s' % a_set.copy())

    buff = []
    for i in range(0, response['totalCount']):
        buff.append(response['webFilters'][i]['info']['ordersNumber'])
    b_set = set(buff)
    logging.info('Номера заказов из Личного кабинета: %s' % b_set.copy())

    result = a_set & b_set
    logging.info('Результат объединения множеств: %s' % result.copy())

    """
    objectId - это и есть id заказа, но есть покупатели, такие как слата
    у которых этот objectId просто так не получить ...
    """
    order_list = []
    for i in range(0, response['totalCount']):
        for ln in result:
            if ln == response['webFilters'][i]['info']['ordersNumber']:
                logging.info('Есть совпадение номеров заказа %s' % ln)

                if response['webFilters'][i]['buyerParty']['gln'] == '4610015769999':  # слата
                    orderId = response['webFilters'][i]['info']['orderId']
                    orderNodeId = response['webFilters'][i]['info']['orderNodeId']
                    driver.get('https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/Order/ToWebObject'
                               '?orderId=' + orderId + '&orderNodeId=' + orderNodeId)
                    try:
                        driver.find_element_by_xpath('//a[@id="NProcess_0"]').click()
                    except NoSuchElementException:
                        logging.error('Ошибка при поиске вкладки /Заявка/, заказчик Слата')
                    objectId = driver.current_url.split('id=')
                    order_list.append(objectId[1])

                if response['webFilters'][i]['buyerParty']['gln'] == '4606068999995':  # лента
                    order_list.append(response['webFilters'][i]['info']['referenceToWebObject']['objectId'])

                if response['webFilters'][i]['buyerParty']['gln'] == '4670014789992':  # командор
                    order_list.append(response['webFilters'][i]['info']['referenceToWebObject']['objectId'])
            #
            # else:
            #     logging.info(data['webFilters'][i]['info']['ordersNumber'])
    logging.info('Список заказов, которые необходимо просмотреть в кабинете:')
    logging.info(order_list)

    # ------------------------------------------------------------------------------------------------------------------
    logging.info('Открываем каждый заказ на анализ План и Факта')
    logging.info('Количество заказов, которые необходимо просмотреть - %s' % str(len(order_list)))
    logging.info('Сначала смотрим План ...')
    list_plan_order = {}
    list_fact_order = {}
    for i in range(0, len(order_list)):
        driver.get('https://edi.kontur.ru/973fa598-96bf-495c-ac98-1ac3a4cd7de2/Supplier/Orders?id=' + order_list[i])
        driver.refresh()
        tree = html.fromstring(driver.page_source)
        gooditem = []
        data_order = tree.xpath('*//div[@id="GoodItemList_GoodItems"]')
        if len(data_order) != 0:
            gooditem.extend(get_gooditem_list(tree))
        else:
            logging.error('Заказ не содержит товаров')
        logging.info(gooditem)

        plan_order = {}
        for j in range(0, len(gooditem)):
            logging.info(gooditem[j][3])
            plan_order[gooditem[j][2]] = gooditem[j][4]
        # pprint.pprint(gooditem)
        # print(plan_order)

        list_plan_order[order_list[i]] = plan_order
        # --------------------------------------------------------------------------------------------------------------
        logging.info('Теперь смотрим Факт ...')
        # делаю обертку, тк вкладка данных по Факту может быть недоступна
        try:
            # обработка вкладки Приемка
            n_process = tree.xpath('//*[contains(@id,"NProcess")]')
            n = len(n_process) - 1
            driver.find_element_by_xpath('//a[@id="NProcess_' + str(n) + '"]').click()
            doc = html.fromstring(driver.page_source)

            data_order = doc.xpath('//div[@class="arrayField__itemData"]')
            gooditem = []

            logging.info('Количество позиций в Факт листе - %s' % len(data_order))
            for j in range(0, len(data_order)):
                gtin = doc.xpath('*//span[@id="GoodItemList_GoodItems_' + str(j) + '_Value_ViewModel_GTIN"]/text()')[0]
                name = doc.xpath('//span[@id="GoodItemList_GoodItems_' + str(j)
                                 + '_Value_ViewModel_Name"]/text()')[0]
                invoic_quantity = doc.xpath('//span[@id="GoodItemList_GoodItems_' + str(j)
                                           + '_Value_ViewModel_InvoicQuantity_CurrentValue"]/text()')[0]
                invoic_quantity = invoic_quantity.replace(' ', '')
                gooditem.append([gtin, invoic_quantity, name])
            logging.info(gooditem)

            fact_order = {}
            for j in range(0, len(gooditem)):
                fact_order[gooditem[j][0]] = gooditem[j][1]
            # print(fact_order)
            list_fact_order[order_list[i]] = fact_order
        except NoSuchElementException:
            # list_fact_order[order_list[i]] =
            logging.error('Счета и документы по данному заказу еще не готовы! Данных по Факту нет')

    print(list_plan_order)
    print(list_fact_order)
    print('-')

    exit()




    # ------------------------------------------------------------------------------------------------------------------
    logging.info('Расчет расхождения между ПЛАНОМ и ФАКТОМ')
    logging.info('work_set %s' % done)
    logging.info('plan_set %s' % plan_set)
    logging.info('fact_set %s' % fact_set)

    # work_set = ['190278', '172383', '165554', '171346', '165558', '171347', '171348', '10005239', '189962', '172203', '', '172202', '172199', '172200', '10013435']
    # plan_set = ['10005239', '165558', '10002903', '10009755', '10005797']
    # fact_set = ['10005239', '10002903', '10009755', '10005797']

    a = set(done)
    b = set(plan_set)
    c = set(fact_set)

    work_plan = a & b
    # for item in work_plan:
    #     logging.info('work_plan % s' % item)
    #     print('plane -', plan[item], 'fact - ', fact[item])
    #     z = plan[item].replace(' ', '')
    #     y = fact[item].replace(' ', '')
    #     z = z.replace(',', '.')
    #     y = y.replace(',', '.')
    #     res = 100 - (float(y)*100/float(z))
    #     print('deviation - ', res)

    #
    # plan_fact = b | c
    # for item in plan_fact:
    #     logging.info('plan_fact % s' % item)

    print('plan - ', plan)
    print('fact - ', fact)
    print('dict - ', dict)

    # вывод итоговой таблицы
    print("{:<20} {:<15} {:<20} {:<10} {:<10} {:<10} {:<10} {:<10} {:<10} {:<10}".format('Date', 'Number', 'Order',
                                                                                         'Product', 'Count', 'Plan',
                                                                                         'Fact', 'Deviation', 'Cost',
                                                                                         'Summa'))
    for k in list:
        buff = k[3].items()
        itemlist = []
        itemlist.extend(iter(buff))
        itemlist.sort()

        header = 0
        dev = summa = 0
        for item in work_plan:

            for j in range(0, len(itemlist)):
                if item == itemlist[j][0] and header == 0:
                    dev = deviation(plan[item], fact[item])
                    # summa = float(itemlist[j][1])*float(cost[item])*float(dev)/float(100)
                    print(
                        "{:<20} {:<15} {:<20} {:<10} {:<10} {:<10} {:<10} {:<10} {:<10} {:<10}".format(k[0], k[1], k[2],
                                                                                                       itemlist[j][0],
                                                                                                       itemlist[j][1],
                                                                                                       plan[item],
                                                                                                       fact[item], dev,
                                                                                                       cost[item],
                                                                                                       summa))
                    header = 1
                elif item == itemlist[j][0] and header == 1:
                    dev = deviation(plan[item], fact[item])
                    # summa = float(itemlist[j][1]) * float(cost[item]) * float(dev) / float(100)
                    print("{:<20} {:<15} {:<20} {:<10} {:<10} {:<10} {:<10} {:<10} {:<10} {:<10}".format('', '', '',
                                                                                                         itemlist[j][0],
                                                                                                         itemlist[j][1],
                                                                                                         plan[item],
                                                                                                         fact[item],
                                                                                                         dev,
                                                                                                         cost[item],
                                                                                                         summa))
        # print("{:<20} {:<15} {:<20} {:<10} {:<10} {:<10}".format(k[0], k[1], k[2], itemlist[0][0], itemlist[0][1], 'pf'))
        # for i in range(1, len(itemlist)):
        #     print("{:<20} {:<15} {:<20} {:<10} {:<10} {:<10}".format('', '', '', itemlist[i][0], itemlist[i][1],'kjh'))

    exit()

    logging.info('report table - %s' % list)

    product = []
    work = []
    plan = []
    fact = []
    for i in range(0, len(list)):
        print(len(list[i][2]))
        for row in list[i][2]:
            print('row - %s' % row)
            for line in work_plan:
                if row == line:
                    product.append(row)
                    work.append(list[i][2].get(row))
    print('report table:')
    for row in list:
        print(row)
    print('report product:')
    for row in product:
        print(row)
    print('work data:')
    for row in work:
        print(row)

    pass


if __name__ == '__main__':
    main()
