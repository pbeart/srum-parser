"""Provides functions to parse define statements from C++ .h files"""

class HeaderFile():
    def __init__(self, path):
        self.file = open(path, "r")
        self.parse_int_defines()

    def parse_int_defines(self):
        self.defines = {}
        for line in self.file.readlines():
            if line.startswith("#define "):
                tokens = line[len("#define "):] # What comes after '#define '

                tokens = tokens.split(" ")

                # Remove empty elements created if there are multiple spaces
                # in a row on a line
                tokens = [token for token in tokens if token != ""]
                
                if len(tokens) > 1: # Isn't the start of a section
                    def_name = tokens[0]
                    
                    def_value = int(tokens[1])

                    self.defines[def_name] = def_value

    def get_name_value(self, name):
        return self.defines.get(name, None)

    def get_value_name(self, value):
        for name in self.defines:
            if self.defines[name] == value:
                return name
        return None

    def __del__(self):
        self.file.close()