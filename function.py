# -*- coding: utf-8 -*-
import csv
import logging
import codecs


def write_csv(fn, data, encode='utf-8'):
    path = 'db/' + fn
    with codecs.open(path, 'w', encoding=encode) as csvFile:
        writef = csv.writer(csvFile, delimiter=';')
        for row in data:
            writef.writerow(row)
    logging.info('Файл успешно создан - %s' % path)
    csvFile.close()


def read_csv(fn, encode='utf-8'):
    data = []
    path = 'db/' + fn
    with codecs.open(path, 'r', encoding=encode) as csvFile:
        readf = csv.reader(csvFile, delimiter=';')
        try:
            for row in readf:
                data.append(row)
            logging.info('Файл успешно прочитан - %s' % path)
        except csv.Error:
            logging.error('Ошибка чтения файла - %s' % path)
    csvFile.close()
    return data
