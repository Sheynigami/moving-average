import logging

from pandas import Series

import settings
from google_spreadsheets import authorize, get_worksheet, input_spreadsheet_id

log = logging.getLogger('moving_average')


class MovingAverage:
    def __init__(self, data, window=2, dtype=int, allow_negative_values=False):
        self._data = data
        self._window = window
        self._dtype = dtype
        self._allow_negative = allow_negative_values

        self._calculate()

    def check_empty_items(self):
        if '' in self._data or None in self._data:
            raise TypeError('Dataset has empty items')

    def check_enough_items(self):
        if len(self._data) < self._window:
            raise Exception('Dataset should have at least %s records' % self._window)

    def check_values_type(self):
        if not all(isinstance(i, self._dtype) for i in self._data):
            raise TypeError('Dataset contains items with illegal type')

    def check_negative_values(self):
        if not all(i >= 0 for i in self._data):
            raise ValueError('Dataset contains negative values')

    def validate_data(self):
        self.check_empty_items()
        self.check_enough_items()
        self.check_values_type()

        if not self._allow_negative:
            self.check_negative_values()

    def _calculate(self):
        self.validate_data()

        self._data = Series(data=self._data, dtype=self._dtype).rolling(window=self._window).mean()

    def get_data(self):
        return self._data


def main():
    client_secret_file = getattr(settings, 'CLIENT_SECRET_FILE', 'client_secret.json')
    client = authorize(client_secret_file)
    if not client:
        return 'Can\'t authenticate to google spreadsheets'

    google_spreadsheet_id = input_spreadsheet_id()
    sheet = get_worksheet(google_spreadsheet_id, client)
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

    moving_average_series = moving_average.get_data()
    try:
        sheet.set_series_column_by_name(moving_average_series, col_name=settings.MOVING_AVERAGE_COL)
    except Exception as error:
        return error

    return 'Moving Average was calculated and stored back to spreadsheet.'


if __name__ == '__main__':
    result = main()
    print(result)
