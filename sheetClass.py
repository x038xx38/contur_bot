import httplib2
import apiclient
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint


class Spreadsheet:
    '''Description'''

    def __init__(self, json_keyfile_name, debugMode=False):
        '''Constructor'''
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        # scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_name, scope)
        self.httpAuth = self.credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build('sheets', 'v4', http=self.httpAuth, cache_discovery=False)
        self.debugMode = debugMode
        self.driveService = None
        self.spreadsheetId = None
        self.sheetId = None
        self.sheetTitle = None
        self.spreadsheetUrl = None
        self.bandedRangeId = None
        self.rowCount = None
        self.colCount = None
        self.requests = []
        self.data = []

    def create_sheets(self, title, sheet, rows, cols=22, locale='ru_RU'):
        '''Create spreadsheet'''
        spreadsheet = self.service.spreadsheets().create(
            body={'properties': {'title': title, 'locale': locale},
                  'sheets': [{'properties': {'sheetType': 'GRID',
                                             'sheetId': 0,
                                             'title': sheet,
                                             'gridProperties': {'rowCount': rows,
                                                                'columnCount': cols}}}]}
        )
        response = spreadsheet.execute()
        if self.debugMode:
            print('def create_sheets():')
            pprint(response)
        self.spreadsheetId = response['spreadsheetId']
        self.sheetId = response['sheets'][0]['properties']['sheetId']
        self.sheetTitle = response['sheets'][0]['properties']['title']
        self.spreadsheetUrl = response['spreadsheetUrl']

    def share_sheets(self, body={'type': 'anyone', 'role': 'reader'}):
        '''
        Share spreadsheet
        body - {'type': 'user', 'role': 'reader/writer', 'emailAddress': email}
        body - {'type': 'anyone', 'role': 'reader/writer'}
        '''
        self.driveService = apiclient.discovery.build('drive', 'v3', http=self.httpAuth, cache_discovery=False)
        shareRes = self.driveService.permissions().create(
            fileId=self.spreadsheetId,
            body=body,
            fields='id'
        )
        response = shareRes.execute()
        if self.debugMode:
            print('def share_sheets():')
            pprint(response)

    def set_spreadsheetById(self, spreadsheetId):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheetId)
        response = spreadsheet.execute()
        if self.debugMode:
            print('def set_spreadsheetById():')
            pprint(response)
        try:
            self.spreadsheetId = response['spreadsheetId']
            self.sheetId = response['sheets'][0]['properties']['sheetId']
            self.sheetTitle = response['sheets'][0]['properties']['title']
            self.spreadsheetUrl = response['spreadsheetUrl']
            self.rowCount = response['sheets'][0]['properties']['gridProperties']['rowCount']
            self.colCount = response['sheets'][0]['properties']['gridProperties']['columnCount']

            # self.bandedRangeId = response['sheets'][0]['bandedRanges'][0]['bandedRangeId']

        except KeyError:
            print('KeyError in set_spreadsheetById')

    def to_grid(self, cells_range):
        """
        "A1:B4" -> {sheetId: 0, startRowIndex: 1, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 3}
        """
        zone = {'sheetId': self.sheetId}
        startCell, endCell = cells_range.split(':')[0:2]
        range_AZ = range(ord('A'), ord('Z') + 1)
        if ord(startCell[0]) in range_AZ:
            zone['startColumnIndex'] = ord(startCell[0]) - ord('A')
            startCell = startCell[1:]
        if ord(endCell[0]) in range_AZ:
            zone['endColumnIndex'] = ord(endCell[0]) - ord('A') + 1
            endCell = endCell[1:]
        if len(startCell) > 0:
            zone['startRowIndex'] = int(startCell) - 1
        if len(endCell) > 0:
            zone['endRowIndex'] = int(endCell)
        if self.debugMode:
            print('def to_grid():')
            pprint(zone)
        return zone

    def appendDimension(self, length, dimension='ROWS'):
        self.requests.append({'appendDimension': {'sheetId': 0,
                                                  'dimension': dimension,
                                                  'length': length}})

    def deleteDimension(self, start, end, dimension='ROWS'):
        self.requests.append({'deleteDimension': {'range': {'sheetId': 0,
                                                            'dimension': dimension,
                                                            'startIndex': start,
                                                            'endIndex': end}}})

    def sizeCells(self, start, end, size, dimension='COLUMNS', sheetId=0):
        self.requests.append({'updateDimensionProperties': {'range': {'sheetId': sheetId,
                                                                      'dimension': dimension,
                                                                      'startIndex': start,
                                                                      'endIndex': end},
                                                            'properties': {'pixelSize': size},
                                                            'fields': 'pixelSize'}})

    def filterColumn(self):
        self.requests.append({
            'setBasicFilter': {
                'filter': {
                    'range': {
                        'sheetId': 1753196879,
                        'startRowIndex': 0,
                        'startColumnIndex': 0,
                    },
                    'criteria': {
                        # 0: {
                        #     'hiddenValues': ['Panel']
                        # },
                        3: {
                            'condition': {
                                'type': 'TEXT_CONTAINS',
                                'values': {
                                    'userEnteredValue': '5645638532'
                                }
                            }
                        }
                    }

                }
            }
        })

    def align_repeatCell(self, zone, horizont, vertical='MIDDLE'):
        '''
            CENTER | LEFT | RIGHT # MIDDLE | TOP  | BOTTOM
        '''
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {'userEnteredFormat': {'horizontalAlignment': horizont,
                                                                            'verticalAlignment': vertical}},
                                             'fields': 'userEnteredFormat'}})

    def header_repeatCell(self, zone, red, green, blue):
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {'userEnteredFormat': {
                                                 # 'backgroundColor': { 'red': red, 'green': green, 'blue': blue },
                                                 'horizontalAlignment': 'CENTER',
                                                 'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
                                                                'bold': True}}},
                                             'fields': 'userEnteredFormat'}})

    def mark_repeatCell(self, zone, red, green, blue):
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {'userEnteredFormat': {
                                                 'backgroundColor': {'red': red, 'green': green, 'blue': blue},
                                                 'horizontalAlignment': 'CENTER'}},
                                             'fields': 'userEnteredFormat'}})

    def odds_repeatCell(self, zone):
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {'userEnteredFormat': {
                                                 'backgroundColor': {'red': 1, 'green': 1, 'blue': 1},
                                                 'horizontalAlignment': 'CENTER',
                                                 'textFormat': {
                                                     'foregroundColor': {'red': 0.4, 'green': 0.4, 'blue': 0.4},
                                                     'fontSize': 9}}},
                                             'fields': 'userEnteredFormat'}})

    def URL_repeatCell(self, zone, horizont='LEFT'):
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {
                                                 # 'userEnteredValue': {'formulaValue': '=HYPERLINK("'+link+'";"'+text+'")'},
                                                 # 'hyperlink':'Google.com'},
                                                 'userEnteredFormat': {
                                                     'horizontalAlignment': horizont,
                                                     'textFormat': {
                                                         'foregroundColor': {'red': 0.4, 'green': 0.4, 'blue': 0.4},
                                                         'fontSize': 9}}},
                                             'fields': 'userEnteredFormat'}})

    def addURL_repeatCell(self, zone, link, text):
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {'userEnteredValue': {
                                                 'formulaValue': '=HYPERLINK("' + link + '";"' + text + '")'},
                                                 'hyperlink': 'Google.com'},
                                             'fields': 'userEnteredValue'}})

    def add_formula(self, cell, formula):
        self.requests.append({'repeatCell': {'range': self.to_grid(cell),
                                             'cell': {'userEnteredValue': {'formulaValue': formula}},
                                             'fields': 'userEnteredValue'}})

    def addBanding(self, zone):
        self.requests.append({'addBanding': {'bandedRange': {'bandedRangeId': 1,
                                                             'range': self.to_grid(zone),
                                                             'rowProperties': {
                                                                 'headerColor': {'red': 0.74, 'green': 0.74,
                                                                                 'blue': 0.74},
                                                                 'firstBandColor': {'red': 0.95, 'green': 0.95,
                                                                                    'blue': 0.95},
                                                                 'secondBandColor': {'red': 1, 'green': 1, 'blue': 1}
                                                             }}}})

    def deleteBanding(self):
        self.requests.append({'deleteBanding': {'bandedRangeId': 1}})

    def mergeCells(self, zone):
        '''
            MERGE_ALL - Создать одно слияние из диапазона
            MERGE_COLUMNS - Создать слияние для каждого столбца в диапазоне
            MERGE_ROWS - Создать слияние для каждой строки в диапазоне
        '''
        self.requests.append({'mergeCells': {'range': self.to_grid(zone), 'mergeType': 'MERGE_ALL'}})

    def update_cells(self, range):
        self.requests.append({
            'updateCells': {
                'fields': '*',
                'range': self.to_grid(range)
            }
        })

    def boldScore_repeatCell(self, zone):
        self.requests.append({'repeatCell': {'range': self.to_grid(zone),
                                             'cell': {'userEnteredFormat': {
                                                 'horizontalAlignment': 'CENTER',
                                                 'textFormat': {'foregroundColor': {'red': 0, 'green': 0, 'blue': 0},
                                                                'bold': True}}},
                                             'fields': 'userEnteredFormat'}})

    def batch_update_spreadsheet(self):
        if self.debugMode:
            print('def batch_update_spreadsheet() requests:')
            pprint(self.requests)
        request = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId,
                                                          body={'requests': self.requests})
        response = request.execute()
        if self.debugMode:
            print('def batch_update_spreadsheet() response:')
            pprint(response)
        self.requests = []

    def addData(self, title, values):
        '''
            For output, if the spreadsheet data is: A1=1,B1=2,A2=3,B2=4, then requesting range=A1:B2,
            majorDimension=ROWS will return [[1,2],[3,4]], where as requesting range=A1:B2,
            majorDimension=COLUMNS will return [[1,3],[2,4]]
        '''
        self.data.append({'range': title,
                          'majorDimension': 'ROWS',
                          'values': values})

    def batch_update_values(self, value_input_option='USER_ENTERED'):
        body = {'value_input_option': value_input_option, 'data': self.data}
        request = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=body)
        response = request.execute()
        if self.debugMode:
            print('def batch_update_values() response:')
            pprint(response)
        self.data = []

    def clear(self, range):
        spreadsheet_id = self.spreadsheetId
        range_ = range
        clear_values_request_body = {}
        request = self.service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=range_,
                                                        body=clear_values_request_body)
        response = request.execute()
        if self.debugMode:
            print('def clear(self, range) response:')
            pprint(response)

    def append(self, range, values):
        spreadsheet_id = self.spreadsheetId
        range_ = range
        value_input_option = 'RAW'
        insert_data_option = 'INSERT_ROWS'
        value_range_body = {'range': range,
                          'majorDimension': 'ROWS',
                          'values': values}
        request = self.service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_,
                                                         valueInputOption=value_input_option,
                                                         insertDataOption=insert_data_option, body=value_range_body)
        response = request.execute()
        if self.debugMode:
            print('def append(self, range):')
            pprint(response)



    def batch_get_values(self, ranges):
        '''
        test
        '''
        valueRanges = []
        spreadsheet_id = self.spreadsheetId
        # ranges = ['Filter']
        value_render_option = 'FORMATTED_VALUE'
        date_time_render_option = 'FORMATTED_STRING'
        major_dimension = 'ROWS'
        request = self.service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges,
                                                                majorDimension=major_dimension,
                                                                valueRenderOption=value_render_option,
                                                                dateTimeRenderOption=date_time_render_option)
        response = request.execute()
        # pprint(response)
        valueRanges = response['valueRanges'][0]['values']
        return valueRanges

    def batch_get_value(self, range):
        spreadsheet_id = self.spreadsheetId
        range_ = range
        value_render_option = 'FORMATTED_VALUE'
        date_time_render_option = 'FORMATTED_STRING'
        # major_dimension = 'ROWS'
        request = self.service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_,
                                                           valueRenderOption=value_render_option,
                                                           dateTimeRenderOption=date_time_render_option)
        response = request.execute()
        pprint(response)
        return response


def get_reportSheet(data, ss):
    '''
    Скелет для таблицы
    '''
    rows = len(data) + 1

    if ss.bandedRangeId == None:
        ss.addBanding('A1:P' + str(rows))
    else:
        ss.deleteBanding()
        ss.addBanding('A1:P' + str(rows))

    if (ss.rowCount != None) and (rows > ss.rowCount):
        dlt = rows - ss.rowCount
        ss.appendDimension(dlt)
    # else:
    #    start = rows
    #    end = ss.rowCount
    #    ss.deleteDimension(start,end)
    ss.sizeCells(0, 1, 25)
    ss.align_repeatCell('B2:B' + str(rows), 'CENTER')
    ss.sizeCells(1, 2, 110)
    ss.align_repeatCell('C2:C' + str(rows), 'LEFT')
    ss.sizeCells(3, 5, 40)
    ss.align_repeatCell('D2:E' + str(rows), 'CENTER')
    ss.align_repeatCell('F2:F' + str(rows), 'LEFT')
    ss.sizeCells(6, 15, 40)
    ss.sizeCells(15, 16, 130)
    ss.mergeCells('K1:L1')
    ss.header_repeatCell('A1:P1', 0.4, 0.4, 0.4)
    ss.align_repeatCell('G2:O' + str(rows), 'CENTER')
    ss.boldScore_repeatCell('K2:L' + str(rows))
    for i in range(2, rows, 2):
        ss.odds_repeatCell('A' + str(i + 1) + ':P' + str(i + 1))
    header = [
        ['', 'Date', 'Team home', 'PWR', 'FRM', 'Team away', 'PWR', 'FRM', 'Index', 'Tip', 'Score', '', 'WH', 'X', 'WA',
         'Url match']]
    ss.addData(ss.sheetTitle, 'A1', header)
    ss.addData(ss.sheetTitle, 'A2', data)
    # j = 3
    # for i in range(1,len(data)):
    #    try:
    #        if float(data[i][12])>2 and float(data[i][13])>2 and float(data[i][14])>2:
    #            ss.addData(ss.sheetTitle,'J'+str(j),[[u'\U00002049']])
    #        j += 2
    #    except ValueError as ex:
    #        print('ValueError - %s'%ex)
    #        j += 2

    result = bt_analysis.find_win(data)
    j = 3
    for i in range(0, len(result)):
        if result[i] == 'green':
            ss.mark_repeatCell('K' + str(j) + ':L' + str(j), 0.93, 0.96, 0.89)
            j += 2
        elif result[i] == 'red':
            ss.mark_repeatCell('K' + str(j) + ':L' + str(j), 0.99, 0.86, 0.90)
            j += 2
        else:
            j += 2

    ss.batch_update_spreadsheet()
    ss.batch_update_values()

    ss.URL_repeatCell('P2:P' + str(rows))
    for i in range(0, len(data), 2):
        link = data[i][15]
        text = data[i][15].split('/')[0]
        ss.addURL_repeatCell('P' + str(i + 2) + ':P' + str(i + 2), link, text)
        ss.URL_repeatCell('P' + str(i + 2) + ':P' + str(i + 2), 'RIGHT')

    ss.batch_update_spreadsheet()

    j = 3
    for i in range(0, len(result)):
        if result[i] == 'green':
            response = ss.batch_get_value('B' + str(j))
            try:
                buff = response['values']
                ss.addData(ss.sheetTitle, 'B' + str(j), [[buff[0][0] + ' ' + u'\U0001F60A']])
            except KeyError:
                ss.addData(ss.sheetTitle, 'B' + str(j), [['']])
            j += 2

        elif result[i] == 'red':
            response = ss.batch_get_value('B' + str(j))
            try:
                buff = response['values']
                ss.addData(ss.sheetTitle, 'B' + str(j), [[buff[0][0] + ' ' + u'\U00002716']])
            except KeyError:
                ss.addData(ss.sheetTitle, 'B' + str(j), [['']])
            j += 2
        else:
            j += 2
    ss.batch_update_values()

    urlPreviewSheet = 'https://docs.google.com/spreadsheets/d/' + ss.spreadsheetId + '/preview'
    return urlPreviewSheet
