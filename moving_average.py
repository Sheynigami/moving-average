import logging

import gspread
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

import settings

log = logging.getLogger('moving_average')


def login_google_spreadsheets(client_secret_filename):
    scope = ['https://spreadsheets.google.com/feeds']
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(client_secret_filename, scope)
        client = gspread.authorize(credentials)
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


def get_spreadsheet(spreadsheet_id, client):
    try:
        book = client.open_by_key(spreadsheet_id)
    except Exception as error:
        log.error('%s' % error)
        return None
    else:
        return book.sheet1


class MovingAverage:
    MOVING_AVERAGE_COL = 'Moving Average'
    VISITORS_COL = 'Visitors'

    def __init__(self, spreadsheet, window=2):
        self._spreadsheet = spreadsheet
        self.window = window

    def visitors_col_exist(self):
        return self.VISITORS_COL in self._spreadsheet.row_values(1)

    def validate_spreadsheet(self):
        if not self.visitors_col_exist():
            raise Exception('Visitors column does not exist.')

    def calculate(self):
        self.validate_spreadsheet()
        dataframe = get_as_dataframe(self._spreadsheet, parse_dates=True)
        dataframe[self.MOVING_AVERAGE_COL] = dataframe.rolling(window=self.window).mean()[self.VISITORS_COL]
        try:
            set_with_dataframe(self._spreadsheet, dataframe)
        except Exception as error:
            log.error('%s' % error)
            return False
        return True


def main():
    client_secret_file = getattr(settings, 'CLIENT_SECRET_FILE', 'client_secret.json')
    client = login_google_spreadsheets(client_secret_file)
    if client:
        google_spreadsheet_id = input_spreadsheet_id()
        sheet = get_spreadsheet(google_spreadsheet_id, client)
        if sheet and sheet.row_count:
            window = getattr(settings, 'MOVING_AVERAGE_WINDOW', 2)
            moving_average = MovingAverage(sheet, window=window)
            result = moving_average.calculate()
            if result:
                print('Moving Average was calculated and stored back to spreadsheet.')
            else:
                print('Something went wrong during calculation.')
        else:
            print('Spreadsheet is empty or does not exist.')
    else:
        print("Can't authenticate to google spreadsheets.")


if __name__ == '__main__':
    main()
