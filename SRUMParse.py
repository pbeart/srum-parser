import pyesedb
import SRUMUtils

def record_as_list(record):
    "Helper for reading pyesedb records as lists of python objects"

    out_list = []

    for column_index in range(record.number_of_values):
        
        value_type = record.get_column_type(column_index)
        raw_value = record.get_value_data(column_index)

        out_list.append(SRUMUtils.parse_ese_value(raw_value, value_type))

    return out_list

class SRUMParser:
    """Class to manage parsing of SRUM .ESE databases.
    Requires given file handle to remain open while in use."""
    def __init__(self, fileObject, registryFolder=None):
        self.esedb = pyesedb.file()
        self.esedb.open_file_object(fileObject)
        self.raw_tables = self.esedb.tables

    def 