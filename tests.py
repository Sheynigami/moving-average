import os
import unittest

import moving_average as ma
import settings


class TestMovingAverage(unittest.TestCase):
    test_secret_file = 'test_client_secret.json'
    test_spreadsheet_id = '1Cc4LKn1tUN3EEU_QpGvKPNIHxiKcBotp30GuVoPScUk'

    def test_spreadsheet_access(self):
        client = ma.login_google_spreadsheets(self.test_secret_file)
        self.assertIsNotNone(client)

        sheet = ma.get_spreadsheet(self.test_spreadsheet_id, client)
        self.assertIsNotNone(sheet)

        sheet = ma.get_spreadsheet('some spreadsheet id that does not exist', client)
        self.assertIsNone(sheet)

    def test_account_from_settings(self):
        secret_filename = settings.CLIENT_SECRET_FILE
        self.assertTrue(os.path.exists(secret_filename), '%s does not exists' % settings.CLIENT_SECRET_FILENAME)

        client = ma.login_google_spreadsheets(secret_filename)
        self.assertIsNotNone(client)

    def test_moving_average_calculation(self):
        client = ma.login_google_spreadsheets(self.test_secret_file)

        sheet = ma.get_spreadsheet(self.test_spreadsheet_id, client)
        if ma.MovingAverage.MOVING_AVERAGE_COL in sheet.row_values(1):
            sheet.resize(cols=3)

        moving_average = ma.MovingAverage(sheet, 2)
        self.assertTrue(moving_average.visitors_col_exist())

        result = moving_average.calculate()
        self.assertTrue(result)
        self.assertIn(ma.MovingAverage.MOVING_AVERAGE_COL, sheet.row_values(1))

        correct_moving_average = ['', '140,758', '144,088', '129,507', '106,907', '122,185']
        self.assertListEqual(correct_moving_average, sheet.col_values(4)[1:])


if __name__ == "__main__":
    unittest.main()
