from xlrd import xldate_as_tuple

from xlwt import Workbook

import datetime as dt
from datetime import date, datetime

from dateutil import parser


def standardize_date(date_object, workbook):
    """
    Grabs the date from excel file and tries to convert it in a datetime object.
    Excel date format and string date format return correct date object (44512, '2013.01.01', '2013-01-01' etc.).
    This is especially useful if you need to make some comparison between date in file and some date.
    """
    if isinstance(date_object, float):
        date_as_tup = xldate_as_tuple(date_object, workbook.datemode)
        return datetime.date(dt.datetime(*date_as_tup))
    else:
        converted_date = parser.parse(date_object)
        return datetime.date(converted_date)
