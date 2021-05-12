"""
xl.py - python library to deal with *big* Excel files.
"""

from abc import ABC
from collections import namedtuple, ChainMap
from contextlib import contextmanager
import datetime
import io
from itertools import zip_longest
import os
import re
import string
import xml.etree.cElementTree as ET
from zipfile import ZipFile

# ISO/IEC 29500:2011 in Part 1, section 18.8.30
STANDARD_STYLES = {
    '0' : 'General',
    '1' : '0',
    '2' : '0.00',
    '3' : '#,##0',
    '4' : '#,##0.00',
    '9' : '0%',
    '10' : '0.00%',
    '11' : '0.00E+00',
    '12' : '# ?/?',
    '13' : '# ??/??',
    '14' : 'mm-dd-yy',
    '15' : 'd-mmm-yy',
    '16' : 'd-mmm',
    '17' : 'mmm-yy',
    '18' : 'h:mm AM/PM',
    '19' : 'h:mm:ss AM/PM',
    '20' : 'h:mm',
    '21' : 'h:mm:ss',
    '22' : 'm/d/yy h:mm',
    '37' : '#,##0 ;(#,##0)',
    '38' : '#,##0 ;[Red](#,##0)',
    '39' : '#,##0.00;(#,##0.00)',
    '40' : '#,##0.00;[Red](#,##0.00)',
    '45' : 'mm:ss',
    '46' : '[h]:mm:ss',
    '47' : 'mmss.0',
    '48' : '##0.0E+0',
    '49' : '@',
}


ExcelErrorValue = namedtuple('ExcelErrorValue', 'value')


class ExcelObj(ABC):
    """
    Abstract base class for other excel objects (workbooks, worksheets, etc.)
    """
    main_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    @staticmethod
    def tag_with_ns(tag, ns):
        "Return XML tag with namespace that can be used with ElementTree"
        return '{%s}%s' % (ns, tag)

    @staticmethod
    def col_num_to_letter(n):
        "Return column letter for column number ``n``"
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    @staticmethod
    def col_letter_to_num(letter):
        "Return column number for column letter ``letter``"
        assert re.match(r'[A-Z]+', letter)
        num = 0
        for char in letter:
            num = num * 26 + (ord(char.upper()) - ord('A')) + 1
        return num


class Worksheet(ExcelObj):
    """
    Excel worksheet
    """

    def __init__(self, workbook, name, number, location=''):
        self._used_area = None
        self._row_length = None
        self._num_rows = None
        self._num_cols = None
        self.workbook = self.wb = workbook
        self.name = name
        self.number = number
        self.location = location or 'xl/worksheets/sheet{number}.xml'

    @contextmanager
    def get_sheet_xml(self):
        "Get a pointer to the xml file underlying the current sheet"
        with self.workbook.xls.open(self.location) as f:
            yield io.TextIOWrapper(f, self.workbook.encoding)

    @property
    def range(self):
        "Return data found in range of cells"
        return Range(self)

    @property
    def rows(self):
        "Iterator that will yield every row in this sheet between start/end"
        return Range(self)

    def _set_dimensions(self):
        "Return the 'standard' row length of each row in this worksheet"
        if ':' not in self.used_area:
            self._num_cols = 0
            self._num_rows = 0
        else:
            _, end = self.used_area.split(':')
            last_col, last_row = re.match(r"([A-Z]+)([0-9]+)", end).groups()
            self._num_cols = self.col_letter_to_num(last_col)
            self._num_rows = int(last_row)

    def _get_num_cols(self):
        "Return the number of standard columns in this worksheet"
        if self._num_cols is None:
            self._set_dimensions()
        return self._num_cols

    def _set_num_cols(self, n):
        "Set the number of columns in the sheet (use with caution!)"
        self._num_cols = n

    num_cols = property(_get_num_cols, _set_num_cols)

    @property
    def num_rows(self):
        "Return the total number of rows used in this worksheet"
        if self._num_rows is None:
            self._set_dimensions()
        return self._num_rows

    @property
    def used_area(self):
        "Return the used area of this sheet"
        if self._used_area is not None:
            return self._used_area
        dimension_tag = self.tag_with_ns('dimension', self.main_ns)
        sheet_data_tag = self.tag_with_ns('sheetData', self.main_ns)
        with self.get_sheet_xml() as sheet:
            for event, elem in ET.iterparse(sheet, events=('start', 'end')):
                if event == 'start':
                    if elem.tag == dimension_tag:
                        used_area = elem.get('ref')
                        if used_area != 'A1':
                            break
                    if elem.tag == sheet_data_tag:
                        # unreliable
                        if list(elem):
                            num_cols = len(list(elem)[0])
                            used_area = f'A1:{num2col(num_cols)}{len(elem)}'
                        break
                elem.clear()
            self._used_area = used_area
        return used_area

    def head(self, num_rows=10):
        "Return first 'num_rows' from this worksheet"
        return self.rows[:num_rows+1] # 1-based

    def cat(self, tab=1):
        "Return/yield all rows from this worksheet"
        dat = self.rows[1] # 1 based!
        XLRec = namedtuple('XLRec', dat[0], rename=True) # pylint: disable=C0103
        for row in self.rows[1:]:
            yield XLRec(*row)


class Range(ExcelObj):
    """
    Excel ranges
    """

    def __init__(self, ws):
        self.worksheet = self.ws = ws
        self.start = None
        self.stop = None
        self.step = None
        self.colstart = None
        self.colstop = None
        self.colstep = None

    def __len__(self):
        return self.worksheet.num_rows

    def __iter__(self):
        with self.ws.get_sheet_xml() as xml_doc:
            row_tag = self.tag_with_ns('row', self.main_ns)
            c_tag = self.tag_with_ns('c', self.main_ns)
            v_tag = self.tag_with_ns('v', self.main_ns)
            row = []
            this_row = -1
            next_row = 1 if self.start is None else self.start
            # last_row = self.ws.num_rows + 1 if self.stop is None else self.stop
            last_row = 1_048_576 if self.stop is None else self.stop
            context = ET.iterparse(xml_doc, events=('start', 'end'))
            context = iter(context)
            event, root = next(context)
            for event, elem in context:
                if event == 'end':
                    if elem.tag == row_tag:
                        this_row = int(elem.get('r'))
                        if this_row >= last_row:
                            break
                        while next_row < this_row:
                            yield self._row([])
                            next_row += 1
                        if this_row == next_row:
                            yield self._row(row)
                            next_row += 1
                        row = []
                        this_row = -1
                        root.clear()
                    elif elem.tag == c_tag:
                        val = elem.findtext(v_tag)
                        if not val:
                            is_elem = elem.find(self.tag_with_ns('is', self.main_ns))
                            if is_elem:
                                val = is_elem.findtext(self.tag_with_ns('t', self.main_ns))
                        if val:
                            # only append cells with values
                            cell = ['', '', '', ''] # ref, type, value, style
                            cell[0] = elem.get('r') # cell ref
                            cell[1] = elem.get('t') # cell type
                            if cell[1] == 's': # string
                                cell[2] = self.ws.workbook.strings[int(val)]
                            else:
                                cell[2] = val
                            cell[3] = elem.get('s') # cell style
                            row.append(cell)

    def __getitem__(self, rng):
        if isinstance(rng, slice):
            if rng.start is not None:
                self.start = rng.start
            if rng.stop is not None:
                self.stop = rng.stop
            if rng.step is not None:
                self.step = rng.step
            matx = [_ for _ in self]
            self.start = self.stop = self.step = None
            return matx
        elif isinstance(rng, str):
            if ':' in rng:
                beg, end = rng.split(':')
            else:
                beg = end = rng
            cell_split = lambda cell: re.match(r"([A-Z]+)([0-9]+)", cell).groups()
            first_col, first_row = cell_split(beg)
            last_col, last_row = cell_split(end)
            first_col = self.col_letter_to_num(first_col) - 1 # python addressing
            first_row = int(first_row)
            last_col = self.col_letter_to_num(last_col)
            last_row = int(last_row)
            self.start = first_row
            self.stop = last_row + 1
            self.colstart = first_col
            self.colstop = last_col
            matx = [_ for _ in self]
            # reset
            self.start = self.stop = self.step = None
            self.colstart = self.colstop = self.colstep = None
            return matx
        elif isinstance(rng, int):
            self.start = rng
            self.stop = rng + 1
            matx = [_ for _ in self]
            self.start = self.stop = self.step = None
            return matx
        else:
            raise NotImplementedError("Cannot understand request")

    def __call__(self, rng):
        return self.__getitem__(rng)

    def _row(self, row):
        lst = [None] * self.ws.num_cols
        col_re = re.compile(r'[A-Z]+')
        col_pos = 0
        for cell in row:
            # apparently, 'r' attribute is optional and some MS products don't
            # spit it out. So we default to incrementing from last known col
            # (or 0 if we are at the beginning) when r is not available.
            if cell[0]:
                col = cell[0][:col_re.match(cell[0]).end()]
                col_pos = self.col_letter_to_num(col) - 1
            else:
                col_pos += 1

            if col_pos >= len(lst):
                # dimensions may not be set right in worksheet
                extend_by = col_pos - len(lst) + 1
                self.ws.num_cols += extend_by
                lst += [None for _ in range(extend_by)]

            try:
                style = self.ws.wb.styles[int(cell[3])]
            except Exception as e:
                style = ''

            # convert to python value (if necessary)
            celltype = cell[1]
            cellvalue = cell[2]
            if celltype in ('str', 's', 'inlineStr'):
                lst[col_pos] = cellvalue
            elif celltype == 'b':
                lst[col_pos] = bool(int(cellvalue))
            elif celltype == 'e':
                lst[col_pos] = ExcelErrorValue(cellvalue)
            elif celltype == 'bl':
                lst[col_pos] = None
            # Lastly, default to a number
            else:
                lst[col_pos] = float(cellvalue)
        colstart = 0 if self.colstart is None else self.colstart
        colstop = self.ws.num_cols if self.colstop is None else self.colstop
        return lst[colstart:colstop]


class Workbook(ExcelObj):
    """
    Excel workbook
    """

    def __init__(self, file_obj, workbook_path=None, encoding='utf8'):
        self.xls = ZipFile(file_obj)
        self.encoding = encoding
        self._strings = None
        self._sheets = None
        self._styles = None
        self.date_system = self.get_date_system()
        if workbook_path:
            self.name = os.path.basename(workbook_path)
            self.path = workbook_path
        else:
            self.name = self.workbook_path = ''

    def get_date_system(self):
        "Determine the date system used by the current workbook"
        with self.xls.open('xl/workbook.xml') as xml_doc:
            tree = ET.parse(io.TextIOWrapper(xml_doc, self.encoding))
            tag = self.tag_with_ns('workbookPr', self.main_ns)
            tag_element = tree.find(tag)
            if tag_element and tag_element.get('date1904') == '1':
                return 1904
            return 1900

    @property
    def sheets(self):
        "Return list of all sheets in workbook"
        if self._sheets is not None:
            return self._sheets
        tag = self.tag_with_ns('sheet', self.main_ns)
        ref_tag = self.tag_with_ns('id', self.rel_ns)
        sheet_map = {}
        locs = {} # locations from relationship id to target location
        with self.xls.open('xl/_rels/workbook.xml.rels') as xml_doc:
            tree = ET.parse(io.TextIOWrapper(xml_doc, self.encoding))
            for rshp in tree.iter(self.tag_with_ns('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')):
                id = rshp.get('Id')
                target = rshp.get('Target')
                locs[id] = target
        with self.xls.open('xl/workbook.xml') as xml_doc:
            tree = ET.parse(io.TextIOWrapper(xml_doc, self.encoding))
            for sheet in tree.iter(tag):
                name = sheet.get('name')
                ref = sheet.get(ref_tag)
                num = int(sheet.get('sheetId'))
                sheet = Worksheet(self, name, num, 'xl/' + locs[ref] if not locs[ref].startswith('/') else locs[ref][1:])
                sheet_map[name] = sheet
                sheet_map[num] = sheet
        self._sheets = sheet_map
        return self._sheets

    @property
    def strings(self):
        "Return list of shared strings within this workbook"
        if self._strings is not None:
            return self._strings
        # Cannot use t element (which we were doing before). See
        # http://bit.ly/2J7xAPu for more info on shared strings.
        tag = self.tag_with_ns('si', self.main_ns)
        strings = []
        with self.xls.open('xl/sharedStrings.xml') as xml_doc:
            tree = ET.parse(io.TextIOWrapper(xml_doc, self.encoding))
            for elem in tree.iter(tag):
                strings.append(''.join(_ for _ in elem.itertext()))
        self._strings = strings
        return strings

    @property
    def styles(self):
        "Return list of styles used within this workbook"
        if self._styles is not None:
            return self._styles
        styles = []
        style_tag = self.tag_with_ns('xf', self.main_ns)
        numfmt_tag = self.tag_with_ns('numFmt', self.main_ns)
        with self.xls.open('xl/styles.xml') as xml_doc:
            tree = ET.parse(io.TextIOWrapper(xml_doc, self.encoding))
            number_fmts_table = tree.find(self.tag_with_ns('numFmts', self.main_ns))
            number_fmts = {}
            if number_fmts_table:
                for num_fmt in number_fmts_table.iter(numfmt_tag):
                    number_fmts[num_fmt.get('numFmtId')] = num_fmt.get('formatCode')
            number_fmts.update(STANDARD_STYLES)
            style_table = tree.find(self.tag_with_ns('cellXfs', self.main_ns))
            if style_table:
                for style in style_table.iter(style_tag):
                    fmtid = style.get('numFmtId')
                    if fmtid in number_fmts:
                        styles.append(number_fmts[fmtid])
        self._styles = styles
        return styles


    def num_to_date(self, number):
        """
        Return date of "number" based on the date system used in this workbook.

        The date system is either the 1904 system or the 1900 system depending
        on which date system the spreadsheet is using. See
        http://bit.ly/2He5HoD for more information on date systems in Excel.
        """
        if self.date_system == 1900:
            # Under the 1900 base system, 1 represents 1/1/1900 (so we start
            # with a base date of 12/31/1899).
            base = datetime.datetime(1899, 12, 31)
            # BUT (!), Excel considers 1900 a leap-year which it is not. As
            # such, it will happily represent 2/29/1900 with the number 60, but
            # we cannot convert that value to a date so we throw an error.
            if number == 60:
                raise ValueError("Bad date in Excel file - 2/29/1900 not valid")
            # Otherwise, if the value is greater than 60 we need to adjust the
            # base date to 12/30/1899 to account for this leap year bug.
            elif number > 60:
                base = base - datetime.timedelta(days=1)
        else:
            # Under the 1904 system, 1 represent 1/2/1904 so we start with a
            # base date of 1/1/1904.
            base = datetime.datetime(1904, 1, 1)
        days = int(number)
        partial_days = number - days
        seconds = int(round(partial_days * 86400000.0))
        seconds, milliseconds = divmod(seconds, 1000)
        if days < -693594:
            return days
        date = base + datetime.timedelta(days, seconds, 0, milliseconds)
        if days == 0:
            return date.time()
        return date


# Some helper functions
def num2col(num):
    """Convert given column letter to an Excel column number."""
    result = []
    while num:
        num, rem = divmod(num-1, 26)
        result[:0] = string.ascii_uppercase[rem]
    return ''.join(result)

def col2num(ltr):
    num = 0
    for c in ltr:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num
