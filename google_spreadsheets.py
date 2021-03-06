import pygsheets
from pandas import np
from pygsheets import Cell, InvalidArgumentValue, Worksheet as BaseWorksheet, format_addr, ValueRenderOption


def input_spreadsheet_id():
    while True:
        spreadsheet_id = input('Google Spreadsheet ID: ')
        if spreadsheet_id:
            return spreadsheet_id


def authorize(client_secret_filename):
    try:
        client = pygsheets.authorize(service_file=client_secret_filename)
    except Exception:
        return None
    else:
        return client


def get_worksheet(spreadsheet_id, client):
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(property='index', value=0)
        worksheet = Worksheet(spreadsheet, worksheet.jsonSheet)
    except Exception:
        return None
    else:
        return worksheet


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

    def get_headers(self):
        return self.get_row(1)

    def get_col(self, col, include_header=False, **kwargs):
        start_row = 0 if include_header else 1
        values = super(Worksheet, self).get_col(col, **kwargs)[start_row:]
        return values

    def get_col_by_name(self, col_name):
        col_index = self.get_col_index(col_name)
        return self.get_col(col_index, include_header=False)

    def get_col_index(self, col_name):
        try:
            index = self.get_headers().index(col_name)
        except ValueError:
            raise Exception('%s column does not exist' % col_name)
        return index + 1

    def set_series_column(self, col_index, series):
        values = series.replace(np.nan, '').tolist()
        values.insert(0, series.name)
        self.update_col(col_index, values)

    def set_series_column_by_name(self, series, col_name=None):
        if col_name is not None:
            series.name = col_name

        try:
            col_index = self.get_col_index(series.name)
        except Exception:
            headers = self.get_headers()
            col_index = len(headers) + 1

        self.set_series_column(col_index, series)
