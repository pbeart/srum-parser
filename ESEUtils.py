import struct

import datetime

def parse_ese_value(raw_value, column_type):
    "Parse a value from an ESE database record into a Python object "\
    "using the given column_type enum. Returns strings/byte arrays in"\
    "their 'raw' form, i.e. bytes if encoding is not known"
    # Value type enums: (python struct module name)
    # 0 - Nil (Invalid column type)
    # 1 - Bool (bool, or None)
    # 2 - Unsigned Byte (bytes of length 1)
    # 3 - Signed 16-bit Int (signed short)
    # 4 - Signed 32-bit Int (int)
    # 5 - 64-bit Currency
    # 6 - 32-bit Float (double)
    # 7 - 64-bit Float
    # 8 - 64-bit Application Time
    # 9 - Binary data
    # 10 - Text (ASCII or Unicode)
    # 11 - Large binary data
    # 12 - Large text (ASCII or Unicode)
    # 13 - Super long value
    # 14 - Unsigned 32-bit Int (unsigned int)
    # 15 - Signed 64-bit Int (long long)
    # 16 - 128-bit GUID
    # 17 - Unsigned 16-bit Int (Unsigned short)

    if raw_value is None: return None

    format_lookup = {
                    1: "?", # Bool
                    2: "B", # Unsigned Byte
                    3: "h", # Signed short
                    4: "i", # Int
                    6: "f", # float
                    7: "d", # double
                    10: "{}s".format(len(raw_value)), # Text
                    12: "{}s".format(len(raw_value)), # Large Text
                    14: "I", # Unsigned Int
                    15: "q", # Signed Long Long Int
                    17: "H" # Unsigned short
                    }

    if column_type in format_lookup:
        try:
            unpacked = struct.unpack("<"+format_lookup[column_type], raw_value)[0]
            
            if column_type in [10,12]: # Text
                try:
                    return unpacked.decode("utf-16")
                except UnicodeDecodeError:
                    return unpacked.decode("utf-8")
                finally:
                    return unpacked
            else:
                return unpacked
        except struct.error as e:
            raise e
            return raw_value

    if column_type == 0: # Nil
        return None

    # 64-bit currency, appears to be equivalent to the datatype documented in:
    # https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/currency-data-type
    elif column_type == 5:  # 64-bit currency
        return None
    elif column_type == 8: # 64-bit application time
        microseconds = struct.unpack("<Q", raw_value)[0] / 10.0 # Interpret as unsigned 64-bit value
        #print(microseconds/(10*1000*1000*60*60*24*365))
        try:
            return datetime.datetime(1601,1,1) + datetime.timedelta(microseconds=microseconds)
        except OverflowError: # Not a valid FILETIME, try parsing as OLE
            as_float = struct.unpack("<d", raw_value)[0]
            return datetime.datetime(1900,1,1) + datetime.timedelta(days=as_float)

    elif column_type == 9: # Binary data
        return raw_value

    elif column_type == 11: # Large binary data
        return raw_value

    elif column_type == 13: # Super long value
        return raw_value # MSDN describes SLV as obsolete

    elif column_type == 16: # 128-bit GUID
        as_hex = raw_value.hex().zfill(32) # 32 chars long, to represent 16 bytes
        return "{"+as_hex[:8]+"-"+as_hex[8:12]+"-"+as_hex[12:14]+"-"+as_hex[14:16]+"-"+as_hex[16:32]+"}"

