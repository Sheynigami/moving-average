import os
import unittest

from pandas import Series, np

import moving_average as ma
import settings
from google_spreadsheets import authorize, get_worksheet


class TestGoogleSpreadsheets(unittest.TestCase):
    test_secret_file = 'test_client_secret.json'
    test_spreadsheet_id = '1Cc4LKn1tUN3EEU_QpGvKPNIHxiKcBotp30GuVoPScUk'

    def setUp(self):
        self.client = authorize(self.test_secret_file)
        self.sheet = get_worksheet(self.test_spreadsheet_id, self.client)

    def test_get_worksheet(self):
        self.assertIsNotNone(self.client)
        self.assertIsNotNone(self.sheet)

        sheet = get_worksheet('some spreadsheet id that does not exist', self.client)
        self.assertIsNone(sheet)

    def test_col_operations(self):
        sheet = get_worksheet(self.test_spreadsheet_id, self.client)

        try:
            sheet.get_col(1)
            sheet.get_col_by_name('Visitors')
        except Exception as error:
            self.fail(error)

        with self.assertRaises(Exception):
            sheet.get_col_by_name('notexistscol')
            sheet.get_col(-1)

        test_list = [1, 2, 3]
        series = Series(test_list, name='test')
        self.sheet.set_series_column_by_name(series)

        col = self.sheet.get_col_by_name('test')
        self.assertListEqual(col, test_list)

    def test_account_access(self):
        secret_filename = settings.CLIENT_SECRET_FILE
        self.assertTrue(os.path.exists(secret_filename), '%s does not exists' % settings.CLIENT_SECRET_FILENAME)

        client = authorize(secret_filename)
        self.assertIsNotNone(client)


class TestMovingAverage(unittest.TestCase):
    def test_validation(self):
        # check illegal characters in list
        with self.assertRaises(TypeError):
            data = [1, 2, 3, 'a']
            ma.MovingAverage(data)

        # check negative values
        with self.assertRaises(ValueError):
            data = [1, 2, -3, 0]
            ma.MovingAverage(data, allow_negative_values=False)

        try:
            data = [1, 2, -3, 0]
            ma.MovingAverage(data, allow_negative_values=True)
        except ValueError:
            self.fail()

        # check empty values in list
        with self.assertRaises(Exception):
            data = [1, 2, 3, '']
            ma.MovingAverage(data)

            data = [1, 2, 3, None]
            ma.MovingAverage(data)

        # check enough elements to calculate moving average
        with self.assertRaises(Exception):
            data = []
            ma.MovingAverage(data)

            data = [1, 2, 3, 4]
            ma.MovingAverage(data, window=5)

    def test_calculation(self):
        test_data = [10, 10, 10, 10]
        moving_average = ma.MovingAverage(test_data).get_data()
        ma_list = moving_average.replace(np.nan, '').tolist()
        result = ['', 10.0, 10.0, 10.0]
        self.assertListEqual(ma_list, result)

        test_data = [0, 10, 5, -4, 8, 16]
        moving_average = ma.MovingAverage(test_data, window=3, allow_negative_values=True).get_data()
        ma_list = moving_average.replace(np.nan, '').tolist()
        result = ['', '', 5.0, 11 / 3, 3.0, 20 / 3]
        self.assertListEqual(ma_list, result)

        test_data = [0, 12, 40, 671, 4, -531]
        moving_average = ma.MovingAverage(test_data, window=6, allow_negative_values=True).get_data()
        ma_list = moving_average.replace(np.nan, '').tolist()
        result = ['', '', '', '', '', sum(test_data) / len(test_data)]
        self.assertListEqual(ma_list, result)


if __name__ == "__main__":
    unittest.main()
