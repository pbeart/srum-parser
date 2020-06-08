"""Module for interacting with ESE databases"""

import pyesedb
import ESEUtils
import parseHeaderDefines
import struct


def SID_bytes_to_string(sid_bytes):
    revision = sid_bytes[0]
    authority_count = sid_bytes[1]

    # Calculate the maximum authority count based on the number of sets of 4
    # bytes of authority information present
    max_authority_count = (len(sid_bytes) - 8) // 4

    # If the given authority count is greater than the maximum count that could
    # be unpacked from the supplied bytes then use the maximum count instead
    authority_count = min(authority_count, max_authority_count)

    # Parse a 48-bit unsigned big-endian value by unpacking it as a 64-bit value
    # with two null bytes at the start
    identifier_authority = struct.unpack(">Q", b"\0\0" + sid_bytes[2:8])[0]

    # Parse identifier_authority number of 4 byte chunks into little endian
    # unsigned ints
    authorities = [
        struct.unpack("<I", sid_bytes[8 + offset * 4:12 + offset * 4])[0]
        for offset in range(authority_count)
    ]

    return "S-{}-{}-{}".format(revision, identifier_authority,
                               "-".join(map(str, authorities)))


def parse_interface_luid(luid_bytes):
    # https://docs.microsoft.com/en-gb/windows/win32/api/ifdef/ns-ifdef-net_luid_lh

    iftypes_hfile = parseHeaderDefines.HeaderFile("headers\\ipifcons.h")

    reserved = luid_bytes[:3]
    netluid_index = luid_bytes[3:6]
    iftype = luid_bytes[6:8]

    iftype_index = struct.unpack("<H", iftype)[0]  # As an int

    iftype_name = iftypes_hfile.get_value_name(iftype_index)

    iftype_nameindex = "{} (#{})".format(iftype_name, iftype_index)

    return {
        "reserved": reserved,
        "netluid_index": netluid_index,
        "iftype": iftype_nameindex
    }


def short_table_name(table_name, processed):
    prefix = "(P) " if processed else ""
    table_aliases = {  # To keep worksheet names under 31 characters
        "{973F5D5C-1D90-4944-BE8E-24B94231A174}": "Network Data Usage Monitor",
        "{D10CA2FE-6FCF-4F6D-848E-B2E99266FA89}": "Application Resource Usage",
        "{DA73FB89-2BEA-4DDC-86B8-6E048C6DA477}": "Energy Estimator ...6DA477}",
        "{DD6636C4-8929-4683-974E-22C046A43763}": "Network Conn. Usage Monitor",
        "{FEE4E14F-02A9-4550-B5CE-5FA2DA202E37}": "Energy Usage",
        "{FEE4E14F-02A9-4550-B5CE-5FA2DA202E37}LT": "Long-term Energy Usage",
        "{D10CA2FE-6FCF-4F6D-848E-B2E99266FA86}": "Push Notifications",
        "{5C8CF1C7-7257-4F13-B223-970EF5939312}": "Energy Estimator ...939312}"
    }
    if table_name in table_aliases:
        return prefix + table_aliases[table_name]
    else:
        return prefix + "Unknown Table ..." + table_name[-7:]


class SRUMParser:
    """Class to manage parsing of SRUM .ESE databases.
    Requires given file handle to remain open while in use."""

    def __init__(self, fileObject, registryFolder=None):
        self.esedb = pyesedb.file()
        self.esedb.open_file_object(fileObject)
        self.raw_tables = self.esedb.tables

    def table_rows(self, table):
        for record in table.records:
            yield self.record_as_list(record)

    def raw_table_rows(self, table):
        for record in table.records:
            yield self.record_as_raw_list(record)

    def record_as_list(self, record):
        "Helper for reading pyesedb records as lists of python objects"

        out_list = []
        for column_index in range(record.number_of_values):

            value_type = record.get_column_type(column_index)
            raw_value = record.get_value_data(column_index)

            out_list.append(ESEUtils.parse_ese_value(raw_value, value_type))

        return out_list

    def record_as_raw_list(self, record):
        "Helper for reading pyesedb records as lists of python objects"

        out_list = []
        for column_index in range(record.number_of_values):
            raw_value = record.get_value_data(column_index)
            out_list.append(raw_value)

        return out_list

    def row_element_by_column_name(self, row, column, table):
        column_names = [
            table.get_column(x).name for x in range(table.number_of_columns)
        ]
        column_index = column_names.index(column)

        return row[column_index]

    def row_index_by_column_name(self, column, table):
        column_names = [
            table.get_column(x).name for x in range(table.number_of_columns)
        ]
        column_index = column_names.index(column)

        return column_index


#print(SID_bytes_to_string(b'\x01\x03\x00\x00\x00\x00\x00\x05Z\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00'))
