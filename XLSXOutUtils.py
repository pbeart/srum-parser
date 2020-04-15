"""Provide helpers for representing data in a way that can be safely stored in an .xlsx file"""

import datetime



def remove_illegal_characters(string):
    """Strip all characters that cannot be stored in an .xlsx file from the
    provided string"""
    string = string.replace("\x00", "")
    return string.encode('unicode_escape').decode('utf-8')

def value_to_safe_string(value):
    """Convert a Python value to a string that can be safely handled in a .xlsx
    file by openpyxl"""
    if isinstance(value, (str, datetime.datetime, float, int)):
        return remove_illegal_characters(value) if isinstance(value, str) else value

    else:
        return repr(value)
