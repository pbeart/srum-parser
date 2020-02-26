import pyesedb
import ESEUtils

import struct

def SID_bytes_to_string(sid_bytes):
    revision = sid_bytes[0]
    authority_count = sid_bytes[1]

    # Parse a 48-bit unsigned big-endian value by unpacking it as a 64-bit value
    # with two null bytes at the start
    identifier_authority = struct.unpack(">Q", b"\0\0"+sid_bytes[2:8])[0]

    # Parse identifier_authority number of 4 byte chunks into little endian
    # unsigned ints
    authorities = [struct.unpack("<I", sid_bytes[8+offset*4:12+offset*4])[0] for offset in range(authority_count)]

    return "S-{}-{}-{}".format(revision, identifier_authority, "-".join(map(str, authorities)))


class SRUMParser:
    """Class to manage parsing of SRUM .ESE databases.
    Requires given file handle to remain open while in use."""
    def __init__(self, fileObject, registryFolder=None):
        self.esedb = pyesedb.file()
        self.esedb.open_file_object(fileObject)
        self.raw_tables = self.esedb.tables

    def table_rows(self, table):
        for record in table.records:
            yield ESEUtils.record_as_list(record)

print(SID_bytes_to_string(b'\x01\x03\x00\x00\x00\x00\x00\x05Z\x00\x00\x00\x00\x00\x00\x00\x01\x00\x00\x00'))