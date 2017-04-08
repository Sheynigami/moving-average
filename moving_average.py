import logging

import pygsheets
from pygsheets import Cell, InvalidArgumentValue, Worksheet as BaseWorksheet, format_addr
from pygsheets import ValueRenderOption

import settings

log = logging.getLogger('moving_average')


class Worksheet(BaseWorksheet):
    def get_values(self, start, end, returnas='matrix', majdim='ROWS', include_empty=True,
                   value_render=ValueRenderOption.UNFORMATTED):
        values = self.client.get_range(self.spreadsheet.id, self._get_range(start, end), majdim.upper(),
                                       value_render=value_render)
        start = format_addr(start, 'tuple')
        if not include_empty:
            matrix = values
        else:
            max_cols = len(max(values, key=len))
            matrix = [list(x + [''] * (max_cols - len(x))) for x in values]

        if returnas == 'matrix':
            return matrix
        else:
            cells = []
            for k in range(len(matrix)):
                row = []
                for i in range(len(matrix[k])):
                    if majdim == 'COLUMNS':
                        row.append(Cell((start[0] + i, start[1] + k), matrix[k][i], self))
                    elif majdim == 'ROWS':
                        row.append(Cell((start[0] + k, start[1] + i), matrix[k][i], self))
                    else:
                        raise InvalidArgumentValue('majdim')

                cells.append(row)
            return cells


def login_google_spreadsheets(client_secret_filename):
    try:
        client = pygsheets.authorize(client_secret_filename)
    except Exception as error:
        log.error('%s' % error)
        return None
    else:
        return client


def input_spreadsheet_id():
    while True:
        spreadsheet_id = input('Google Spreadsheet ID: ')
        if spreadsheet_id:
            return spreadsheet_id


def get_worksheet(spreadsheet_id, client):
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(property='index', value=0)
        worksheet = Worksheet(spreadsheet, worksheet.jsonSheet)
    except Exception as error:
        log.error('%s' % error)
        return None
    else:
        return worksheet


def column_exists(column, spreadsheet):
    return column in spreadsheet.row_values(1)


class MovingAverage:
    MOVING_AVERAGE_COL = 'Moving Average'

    def __init__(self, spreadsheet, col_name='Visitors', col_type=int, window=2):
        self.column = col_name
        self._spreadsheet = spreadsheet
        self.window = window
        self.col_type = col_type

    def get_column(self, col, include_header=False):
        start_row = 0 if include_header else 1
        dataset = self._spreadsheet.get_col(col)[start_row:]
        return dataset

    def get_column_by_name(self, col_name):
        col_index = self.get_col_index(col_name)
        return self.get_column(col_index, include_header=False)

    def get_col_index(self, col_name):
        self.validate_col_exists(col_name)
        return self._spreadsheet.get_row(1).index(col_name) + 1

    def validate_col_exists(self, col):
        if col not in self._spreadsheet.get_row(1):
            raise Exception('%s column does not exist.' % self.column)

    def validate_illegal_characters(self, dataset):
        raise TypeError('%s column contain illegal characters.' % self.column)

    def validate_enough_elements(self, dataset):
        if len(dataset) < self.window:
            raise Exception('%s column should have at least %s records.' % (self.column, self.window))

    def validate_dataset(self, dataset):
        self.validate_enough_elements(dataset)
        # self.validate_illegal_characters(dataset)
        # if not self.check_illegal_characters():
        #     raise TypeError('%s column contain illegal characters.' % dataset)

    def calculate(self):
        dataset = self.get_column_by_name(self.column)
        print(dataset)
        self.validate_dataset(dataset)

        # dataframe = get_as_dataframe(self._spreadsheet, parse_dates=True)
        # dataframe[self.MOVING_AVERAGE_COL] = dataframe.rolling(window=self.window).mean()[self.column]
        # try:
        #     set_with_dataframe(self._spreadsheet, dataframe)
        # except Exception as error:
        #     log.error('%s' % error)
        #     return False
        # return True


VISITORS_COL = 'Visitors'
MOVING_AVERAGE_COL = 'Moving Average'


def main():
    client_secret_file = getattr(settings, 'CLIENT_SECRET_FILE', 'client_secret.json')
    client = login_google_spreadsheets(client_secret_file)
    if not client:
        return 'Can\'t authenticate to google spreadsheets.'

    # google_spreadsheet_id = input_spreadsheet_id()
    # sheet = get_spreadsheet(google_spreadsheet_id, client)
    sheet = get_worksheet('1Cc4LKn1tUN3EEU_QpGvKPNIHxiKcBotp30GuVoPScUk', client)
    if not (sheet and sheet.rows):
        return 'Spreadsheet is empty or does not exist.'

    # window = getattr(settings, 'MOVING_AVERAGE_WINDOW', 2)
    window = 10
    moving_average = MovingAverage(sheet, window=window)
    try:
        result = moving_average.calculate()
    except Exception as error:
        return '%s' % error

    return 'Moving Average was calculated and stored back to spreadsheet.'


if __name__ == '__main__':
    result = main()
    print(result)
