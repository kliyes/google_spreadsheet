# encoding=utf8
'''
Created on 2016-11-10

@author: jingyang <jingyang@nexa-corp.com>
'''
import re

from googleapiclient.errors import HttpError

from google_spreadsheet import exceptions


class Dimension(object):
    """
    Doc: https://developers.google.com/sheets/reference/rest/v4/spreadsheets.values#dimension
    """
    ROWS = "ROWS"
    COLUMNS = "COLUMNS"


class ValueInputOption(object):
    """
    Doc: https://developers.google.com/sheets/reference/rest/v4/ValueInputOption
    """
    RAW = "RAW"
    USER_ENTERED = "USER_ENTERED"


class NumberFormatType(object):
    """
    Doc: https://developers.google.com/sheets/reference/rest/v4/spreadsheets#NumberFormatType
    """
    TEXT = "TEXT"
    NUMBER = "NUMBER"
    PERCENT = "PERCENT"
    CURRENCY = "CURRENCY"
    DATE = "DATE"
    TIME = "TIME"
    DATE_TIME = "DATE_TIME"
    SCIENTIFIC = "SCIENTIFIC"


class Client(object):
    """
    An instance of this class communicates with Google Spreadsheets APIs.

    :param service: client service: utils.spreadsheet_service(...)
    """
    def __init__(self, service):
        self.service = service

    def open(self, file_id):
        """
        Opens a spreadsheet specified by `file id`

        :param file_id: An id of a spreadsheet as it appears in a URL in a browser
        :return: Spreadsheet object
        """
        try:
            response = self.service.spreadsheets().get(spreadsheetId=file_id).execute()
        except HttpError as error:
            if error.resp.status == 404:
                raise exceptions.NotFound(error)
            elif error.resp.status == 403:
                raise exceptions.PermissionDenied(error)
            else:
                raise exceptions.APIException(error)
        else:
            return Spreadsheet(self, response)

    def update(self, file_id, requests):
        try:
            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=file_id,
                body={
                    "requests": requests
                }
            ).execute()
        except HttpError as error:
            if error.resp.status == 400:
                raise exceptions.BadRequest(error)
        else:
            return response

    def values_update(self, file_id, range_name, values, value_input_option=ValueInputOption.USER_ENTERED,
                      major_dimension=Dimension.ROWS):
        try:
            response = self.service.spreadsheets().values().update(
                spreadsheetId=file_id, range=range_name, valueInputOption=value_input_option,
                body={
                    "values": values,
                    "majorDimension": major_dimension
                }
            ).execute()
        except HttpError as error:
            raise exceptions.BadRequest(error)
        else:
            return response

    def values_batch_update(self, file_id, data, value_input_option=ValueInputOption.USER_ENTERED):
        try:
            response = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=file_id,
                body={
                    "valueInputOption": value_input_option,
                    "data": data
                }
            ).execute()
        except HttpError as error:
            raise exceptions.BadRequest(error)
        else:
            return response


class Spreadsheet(object):
    """
    A class for a spreadsheet object
    """
    def __init__(self, client, details):
        self.client = client
        self.details = details

    @property
    def file_id(self):
        return self.details["spreadsheetId"]

    @property
    def title(self):
        return self.details["properties"]["title"]

    def refresh(self):
        """
        Refresh/Re-open this spreadsheet file

        :return: None
        """
        self.details = self.client.open(self.file_id).details

    def all_sheets(self, include_hidden=False):
        """
        Get all sheets list of this spreadsheet

        :param include_hidden: hidden sheets included or not
        :return: list of Sheet objects
        """
        sheets = self.details["sheets"]
        if not include_hidden:
            sheets = filter(lambda x: "hidden" not in x["properties"], sheets)
        return sheets

    def find_sheet_by(self, by, value, include_hidden=False):
        """
        Get a sheet with given value by id, name or index

        :param by: find sheet by, id, name or index
        :param value: find by value
        :param include_hidden: hidden sheets included or not
        :return: Sheet object
        """
        if by == "id":
            key = "sheetId"
        elif by == "name":
            key = "title"
        elif by == "index":
            key = "index"
        else:
            raise

        sheets = self.all_sheets(include_hidden)
        try:
            if by == "index":
                sheet = sheets[value]
            else:
                sheet = (ws for ws in sheets if ws["properties"][key] == value).next()
        except (StopIteration, IndexError):
            raise exceptions.NotFound("Sheet not found")
        else:
            return Sheet(self, sheet)

    def find_sheet_by_id(self, sheet_id, include_hidden=False):
        """
        Shortcut for find_sheet_by(id)

        :param sheet_id: sheet id, integer
        :param include_hidden: hidden sheets included or not
        :return: Sheet object
        """
        return self.find_sheet_by("id", sheet_id, include_hidden)

    def find_sheet_by_name(self, sheet_name, include_hidden=False):
        """
        Shortcut for find_sheet_by(name)

        :param sheet_name: sheet name
        :param include_hidden: hidden sheets included or not
        :return: Sheet object
        """
        return self.find_sheet_by("name", sheet_name, include_hidden)

    def find_sheet_by_index(self, index, include_hidden=False):
        """
        Shortcut for find_sheet_by(index)

        :param index: 0-based index of sheets list
        :param include_hidden: hidden sheets included or not
        :return: Sheet object
        """
        return self.find_sheet_by("index", index, include_hidden)

    def batch_update(self, requests):
        """
        Batch update requests

        :param requests: update requests
        :return: update response
        """
        return self.client.update(self.file_id, requests)

    def add_sheet_request(self, sheet_name, row_count=1000, col_count=1000):
        """
        Only get `add_sheet` request, for batch_update

        :param sheet_name: new sheet name
        :param row_count: new sheet row count
        :param col_count: new sheet column count
        :return: update request
        """
        request = {
            "addSheet": {
                "properties": {
                    "title": sheet_name,
                    "gridProperties": {
                        "rowCount": row_count,
                        "columnCount": col_count
                    }
                }
            }
        }
        return request

    def add_sheet(self, sheet_name, row_count=1000, col_count=1000):
        """
        Add a sheet to this spreadsheet

        :param sheet_name: new sheet name
        :param row_count: new sheet row count
        :param col_count: new sheet column count
        :return: Sheet object
        """
        requests = [self.add_sheet_request(sheet_name, row_count, col_count)]
        response = self.client.update(self.file_id, requests)
        self.refresh()
        return Sheet(self, response["replies"][0]["addSheet"])

    def delete_sheet_request(self, sheet_id):
        """
        Only get `delete_sheet` request, for batch_update

        :param sheet_id: id of a sheet to be deleted
        :return: update request
        """
        request = {
            "deleteSheet": {
                "sheetId": sheet_id
            }
        }
        return request

    def delete_sheet(self, sheet_id):
        """
        Delete the sheet with given id

        :param sheet_id: id of a sheet to be deleted
        :return: None
        """
        requests = [self.delete_sheet_request(sheet_id)]
        self.client.update(self.file_id, requests)
        self.refresh()

    def change_title_request(self, new_title):
        """
        Only get `change_title` request, for batch_update

        :param new_title: new title
        :return: update request
        """
        request = {
            "updateSpreadsheetProperties": {
                "properties": {
                    "title": new_title
                },
                "fields": "title"
            }
        }
        return request

    def change_title(self, new_title):
        """
        Change spreadsheet title

        :param new_title: new title
        :return: None
        """
        requests = [self.change_title_request(new_title)]
        self.client.update(self.file_id, requests)
        self.refresh()


class Sheet(object):
    """
    A class for a sheet/tab in a spreadsheet
    """
    def __init__(self, spreadsheet, details):
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client
        self.details = details

    @property
    def sheet_id(self):
        return self.details["properties"]["sheetId"]

    @property
    def name(self):
        return self.details["properties"]["title"]

    @property
    def row_count(self):
        return self.details["properties"]["gridProperties"]["rowCount"]

    @property
    def col_count(self):
        return self.details["properties"]["gridProperties"]["columnCount"]

    _MAGIC_NUMBER = 64
    _cell_addr_re = re.compile(r'([A-Za-z]+)([1-9]\d*)')

    def get_int_addr(self, label):
        """
        Translates cell's label address to a tuple of integers.
        The result is a tuple containing `row` and `column` numbers.

        :param label: String with cell label in common format, e.g. 'B1'.
                      Letter case is ignored.
        Example:
        >>> sheet.get_int_addr('A1')
        >>> (0, 0)
        """
        m = self._cell_addr_re.match(label)
        if m:
            column_label = m.group(1).upper()
            row = int(m.group(2)) - 1

            col = -1
            for i, c in enumerate(reversed(column_label)):
                col += (ord(c) - self._MAGIC_NUMBER) * (26 ** i)
        else:
            raise exceptions.IncorrectCellLabel(label)

        return row, col

    def get_addr_int(self, row, col):
        """
        Translates cell's tuple of integers to a cell label.
        The result is a string containing the cell's coordinates in label form.

        :param row: The row of the cell to be converted.
                    Rows start at index 0.
        :param col: The column of the cell to be converted.
                    Columns start at index 0.
        Example:
        >>> sheet.get_addr_int(0, 0)
        >>> A1
        """
        row = int(row)
        col = int(col)

        if row < 0 or col < 0:
            raise exceptions.IncorrectCellLabel('(%s, %s)' % (row, col))

        div = col + 1
        column_label = ''

        while div:
            (div, mod) = divmod(div, 26)
            if mod == 0:
                mod = 26
                div -= 1
            column_label = chr(mod + self._MAGIC_NUMBER) + column_label

        label = '%s%s' % (column_label, row + 1)
        return label

    def refresh(self):
        self.spreadsheet.refresh()
        sheet = self.spreadsheet.find_sheet_by_id(self.sheet_id)
        self.details = sheet.details
        self.spreadsheet = sheet.spreadsheet

    def delete(self):
        """
        Delete sheet itself

        :return: None
        """
        return self.spreadsheet.delete_sheet(self.sheet_id)

    def update_properties_request(self, properties, fields):
        """
        Get request for updating sheet properties

        :param properties: properties to be updated, dict type
        :param fields: fields to be updated
        :return: update request
        """
        request = {
            "updateSheetProperties": {
                "properties": properties,
                "fields": fields
            }
        }
        return request

    def change_name_request(self, new_name):
        """
        Only get `change_name` request, for batch_update

        :param new_name: name to change
        :return: update request
        """
        request = self.update_properties_request(
            properties={
                "sheetId": self.sheet_id,
                "title": new_name
            },
            fields="title"
        )
        return request

    def change_name(self, new_name):
        """
        Change sheet name

        :param new_name: name to change
        :return: None
        """
        requests = [self.change_name_request(new_name)]
        self.client.update(self.spreadsheet.file_id, requests)
        self.refresh()

    def resize_request(self, row=None, col=None):
        """
        Only get `resize` request, for batch_update

        :param row: new row
        :param col: new column
        :return: update request
        """
        request = self.update_properties_request(
            properties={
                "sheetId": self.sheet_id,
                "gridProperties": {
                    "rowCount": row,
                    "columnCount": col
                }
            },
            fields="{}{}".format(
                "gridProperties.rowCount," if row is not None else "",
                "gridProperties.columnCount," if col is not None else ""
            )
        )
        return request

    def resize(self, row=None, col=None):
        """
        Resize sheet

        :param row: new row count
        :param col: new column count
        :return: None
        """
        requests = [self.resize_request(row, col)]
        self.client.update(self.spreadsheet.file_id, requests)
        self.refresh()

    def append_request(self, dimension, length=1):
        """
        Only get `append` request, for batch_update

        :param dimension: ROWS or COLUMNS
        :param length: rows/columns count to be appended
        :return: update request
        """
        request = {
            "appendDimension": {
                "sheetId": self.sheet_id,
                "dimension": dimension,
                "length": length
            }
        }
        return request

    def append(self, row=None, col=None):
        """
        Append empty rows and/or columns

        :param row: rows count to append
        :param col: columns count to append
        :return: None
        """
        requests = []
        if row is not None:
            requests.append(self.append_request(Dimension.ROWS, row))
        if col is not None:
            requests.append(self.append_request(Dimension.COLUMNS, col))

        self.client.update(self.spreadsheet.file_id, requests)
        self.refresh()

    def insert_request(self, dimension, start_index=0, length=1, inherit_before=True):
        """
        Only get `insert` request, for batch_update

        :param dimension: ROWS or COLUMNS
        :param start_index: 0-based index
        :param length: numbers of `dimension` to be inserted
        :param inherit_before: inheritFromBefore
        :return: update request
        """
        if start_index == 0:
            inherit_before = False
        request = {
            "insertDimension": {
                "range": {
                    "sheetId": self.sheet_id,
                    "dimension": dimension,
                    "startIndex": start_index,
                    "endIndex": start_index + length
                },
                "inheritFromBefore": inherit_before
            }
        }
        return request

    def insert(self, row=None, row_start_index=0, col=None, col_start_index=0, inherit_before=True):
        """
        Insert empty rows and/or columns

        :param row: rows count to insert
        :param row_start_index: insert start at row index
        :param col: columns count to insert
        :param col_start_index: insert start at column index
        :param inherit_before: True - give the new columns or rows the same properties as the prior row or column
        :return: None
        """
        requests = []
        if row is not None:
            requests.append(self.insert_request(Dimension.ROWS, row_start_index, row, inherit_before))
        if col is not None:
            requests.append(self.insert_request(Dimension.COLUMNS, col_start_index, col, inherit_before))

        self.client.update(self.spreadsheet.file_id, requests)
        self.refresh()

    def number_format(self, format_type=NumberFormatType.NUMBER, pattern=None):
        """
        Number format object

        :param format_type: format type
        :param pattern: pattern
        :return: number format object
        """
        number_format = {
            "numberFormat": {
                "type": format_type
            }
        }
        if pattern is not None:
            number_format["number_format"]["pattern"] = pattern

        return number_format

    def text_format(self, color=None, font_family=None, font_size=None, bold=None, italic=None,
                    strikethrough=None, underline=None):
        """
        Text format object

        :param color: Color object, eg: {"red": 0, "green": 1, "blue": 1, "alpha": 0}
        :param font_family: string
        :param font_size: size of font
        :param bold: true if the text is bold
        :param italic: true if the text is italic
        :param strikethrough: true if the text has a strikethrough
        :param underline: true if the text is underlined
        :return: text format object
        """
        text_format = {
            "textFormat": {}
        }
        if color is not None:
            text_format["textFormat"]["foregroundColor"] = color
        if font_family is not None:
            text_format["textFormat"]["fontFamily"] = font_family
        if font_size is not None:
            text_format["textFormat"]["fontSize"] = font_size
        if bold is not None:
            text_format["textFormat"]["bold"] = bold
        if italic is not None:
            text_format["textFormat"]["italic"] = italic
        if strikethrough is not None:
            text_format["textFormat"]["strikethrough"] = strikethrough
        if underline is not None:
            text_format["textFormat"]["underline"] = underline

        return text_format

    def format_range_request(self, format_body,
                             start_row_index=None, end_row_index=None,
                             start_col_index=None, end_col_index=None, fields=None):
        """
        Only get `format_range` request, for batch_update

        :param format_body: format body, self.number_format()
        :param start_row_index: 0-based row index
        :param end_row_index: 0-base row index, exclude
        :param start_col_index: 0-base column index
        :param end_col_index: 0-base column index, exclude
        :param fields: fields to update
        :return: update request
        """
        request = {
            "repeatCell": {
                "range": {
                    "sheetId": self.sheet_id
                },
                "cell": {
                    "userEnteredFormat": format_body
                }
            }
        }

        if start_row_index is not None:
            request["repeatCell"]["range"]["startRowIndex"] = start_row_index
        if end_row_index is not None:
            request["repeatCell"]["range"]["endRowIndex"] = end_row_index
        if start_col_index is not None:
            request["repeatCell"]["range"]["startColumnIndex"] = start_col_index
        if end_col_index is not None:
            request["repeatCell"]["range"]["endColumnIndex"] = end_col_index
        if fields is None:
            fields_keys = request["repeatCell"]["cell"]["userEnteredFormat"].keys()
            fields = ",".join(["userEnteredFormat.{}".format(f) for f in fields_keys])

        request["repeatCell"]["fields"] = fields
        return request

    def format_number(self, start_row_index, end_row_index, start_col_index, end_col_index,
                      format_type=NumberFormatType.NUMBER, pattern=None):
        """
        Format cell values for numbers

        :param start_row_index: 0-based row index
        :param end_row_index: 0-base row index, exclude
        :param start_col_index: 0-based column index
        :param end_col_index: 0-based column index, exclude
        :param format_type: number format type
        :param pattern: pattern
        :return: None
        """
        number_format = self.number_format(format_type, pattern)
        requests = [self.format_range_request(number_format, start_row_index, end_row_index,
                                              start_col_index, end_col_index)]
        self.client.update(self.spreadsheet.file_id, requests)
        self.refresh()

    def update_values_data(self, range_start, range_end=None, values=None, major_dimension=Dimension.ROWS):
        """
        Generate update values data, dict type, for batch update values

        :param range_start: cell range start, e.g: A1
        :param range_end: cell range end, e.g: C5
        :param values: values list, e.g: [["A1", "B1"], ["A2", "B2"]]
        :param major_dimension: Indicates which dimension an operation should apply to
        :return: values data
        """
        range_name = "{sheet_name}!{range_start}{range_end}".format(
            sheet_name=self.name, range_start=range_start,
            range_end=":{range_end}".format(range_end=range_end) if range_end else ""
        )
        if not values:
            return {}

        data = {
            "range": range_name,
            "majorDimension": major_dimension,
            "values": values
        }
        return data

    def update_values(self, range_start, range_end=None, values=None,
                      value_input_option=ValueInputOption.USER_ENTERED, major_dimension=Dimension.ROWS):
        """
        Update a single range cells values

        :param range_start: cell range start
        :param range_end: cell range end
        :param values: values list
        :param value_input_option: value input option
        :param major_dimension: major dimension
        :return: updated response
        """
        data = self.update_values_data(range_start, range_end, values, major_dimension)
        if data:
            response = self.client.values_update(
                self.spreadsheet.file_id, data["range"], data["values"], value_input_option, major_dimension)
            self.refresh()
            return response

    def batch_update_values(self, values_data_list, value_input_option=ValueInputOption.USER_ENTERED):
        """
        Update multiple range cells values

        :param values_data_list: values data list, generated from self.update_values_data()
        :param value_input_option: value input option
        :return: updated response
        """
        response = self.client.values_batch_update(self.spreadsheet.file_id, values_data_list, value_input_option)
        self.refresh()
        return response
