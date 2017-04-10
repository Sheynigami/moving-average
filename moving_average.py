import logging

from pandas import Series

import settings
from google_spreadsheets import authorize, get_worksheet

log = logging.getLogger('moving_average')


class MovingAverage:
    def __init__(self, data, window=2, dtype=int):
        self._data = data
        self._window = window
        self._dtype = dtype

        self._calculate()

    def check_empty_items(self):
        if '' in self._data or None in self._data:
            raise TypeError('Dataset has empty items')

    def check_enough_items(self):
        if len(self._data) < self._window:
            raise Exception('Dataset should have at least %s records' % self._window)

    def validate_data(self):
        self.check_empty_items()
        self.check_enough_items()

    def _calculate(self):
        self.validate_data()

        self._data = Series(data=self._data, dtype=self._dtype).rolling(window=self._window, min_periods=2).mean()

    def get_data(self):
        return self._data


def main():
    client_secret_file = getattr(settings, 'CLIENT_SECRET_FILE', 'client_secret.json')
    client = authorize(client_secret_file)
    if not client:
        return 'Can\'t authenticate to google spreadsheets'

    # google_spreadsheet_id = input_spreadsheet_id()
    # sheet = get_spreadsheet(google_spreadsheet_id, client)
    sheet = get_worksheet('1Cc4LKn1tUN3EEU_QpGvKPNIHxiKcBotp30GuVoPScUk', client)
    if not (sheet and sheet.rows):
        return 'Spreadsheet is empty or does not exist'

    window = getattr(settings, 'MOVING_AVERAGE_WINDOW', 2)
    try:
        visitors = sheet.get_col_by_name(settings.VISITORS_COL)
    except Exception as error:
        return error

    try:
        moving_average = MovingAverage(visitors, window=window)
    except Exception as error:
        return '%s' % error

    moving_average = moving_average.get_data()
    try:
        sheet.set_series_column_by_name(moving_average, col_name=settings.MOVING_AVERAGE_COL)
    except Exception as error:
        return error

    return 'Moving Average was calculated and stored back to spreadsheet.'


if __name__ == '__main__':
    result = main()
    print(result)
