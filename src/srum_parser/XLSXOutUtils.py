"""Provide helpers for representing data in a way that can be safely stored in an .xlsx file,
and displaying data in a user-friendly format."""

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
        return remove_illegal_characters(value) if isinstance(value,
                                                              str) else value

    else:
        return repr(value)


def num_secs_display(num_secs):
    """Convert a number of seconds toa nicely formatted string, e.g. 3601 > 1h0m1s.
    num_secs can be a float or an int."""

    negative = "-" if num_secs < 0 else ""
    num_secs = abs(num_secs)

    secs = num_secs % 60
    mins = (num_secs // 60) % 60
    hours = (num_secs // 3600)

    secs = int(secs) if int(secs) == secs else secs
    mins = int(mins) if int(mins) == mins else mins
    hours = int(hours) if int(hours) == hours else hours

    if hours <= 0:
        if mins <= 0:  # Only need to display seconds
            return "{}{:02}s".format(negative, secs)
        else:  # Display minutes, seconds
            return "{}{:02}m{:02}s".format(negative, mins, secs)
    else:  # Display hours, minutes, seconds
        return "{}{:02}h{:02}m{:02}s".format(negative, hours, mins, secs)


def num_bytes_display(num_bytes):
    "Convert a number of bytes to a nicely formatted string, e.g. 1100 > 1.1kb"
    display_format = min(len(str(num_bytes)) // 3, 6)

    display_suffixes = {
        0: "b",
        1: "Kb",
        2: "Mb",
        3: "Gb",
        4: "Tb",
        5: "Pb",
        6: "Eb"
    }

    display_suffix = display_suffixes[display_format]

    shown_value = num_bytes / (10**(display_format * 3))

    shown_value = int(shown_value) if int(
        shown_value) == shown_value else shown_value

    return "{}{}".format(shown_value, display_suffix)
