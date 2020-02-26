import ESEUtils
import pytest
import datetime

class TestParseEseValues:
    def test_nil(self):
        assert ESEUtils.parse_ese_value(b"teststring", 0) == None

    def test_bool(self):
        assert ESEUtils.parse_ese_value(b"\x11", 1) == True
        assert ESEUtils.parse_ese_value(b"\x00", 1) == False

    def test_byte(self):
        assert ESEUtils.parse_ese_value(b"\x05", 2) == 5
        assert ESEUtils.parse_ese_value(b"\xff", 2) == 255

    def test_signed_short(self):
        assert ESEUtils.parse_ese_value(b"\xff\xff", 3) == -1
        assert ESEUtils.parse_ese_value(b"\x45\x03", 3) == 0x0345
    
    def test_signed_int(self):
        assert ESEUtils.parse_ese_value(b"\xff\xff\xff\xff", 4) == -1
        assert ESEUtils.parse_ese_value(b"\x01\x45\x03\x00", 4) == 0x034501

    def test_float(self):
        assert ESEUtils.parse_ese_value(b'\x00\x00\x80\xbf', 6) == -1
        assert ESEUtils.parse_ese_value(b'@@QH', 6) == 0x034501

    def test_double(self):
        assert ESEUtils.parse_ese_value(b'\x00\x00\x00\x00\x00\x00\xf0\xbf', 7) == -1
        assert ESEUtils.parse_ese_value(b"\x00\x00\x00\x00\x08\x28\x0a\x41", 7) == 0x034501
    
    def test_datetime(self):
        # FILETIME:
        target_date = datetime.datetime(2010, 6, 29, 9, 47, 42, 754212)
        parsed_date = ESEUtils.parse_ese_value(b'\x01\xcb\x17\x70\x1e\x9c\x88\x5a'[::-1], 8)
        assert abs((target_date - parsed_date).total_seconds()) < 1

        # OLE Time:
        target_date = datetime.datetime(2019, 12, 22, 13, 18,)
        parsed_date = ESEUtils.parse_ese_value(b'\xbc\xbb\xbb\xbbqe\xe5@', 8)
        assert abs((target_date - parsed_date).total_seconds()) < 1

    def test_binary_largebinary(self):
        assert ESEUtils.parse_ese_value(b'\xde\xad\xbe\xef', 9) == repr(b'\xde\xad\xbe\xef') # binary
        assert ESEUtils.parse_ese_value(b'\xde\xad\xbe\xef', 11) == repr(b'\xde\xad\xbe\xef') # large binary

    def test_text_largetext(self):
        assert ESEUtils.parse_ese_value(b'ham sandwiches', 10) == b"ham sandwiches" # text
        assert ESEUtils.parse_ese_value(b'sam handwiches', 12) == b"sam handwiches" # large text

    def test_unsigned_int(self):
        assert ESEUtils.parse_ese_value(b"\xff\xff\xff\xff", 14) == 4294967295
        assert ESEUtils.parse_ese_value(b"\x01\x45\x03\x00", 14) == 0x034501

    def test_signed_long_long(self):
        assert ESEUtils.parse_ese_value(b"\xff\xff\xff\xff\xff\xff\xff\xff", 15) == -1
        assert ESEUtils.parse_ese_value(b"\x01\x45\x03\x00\x00\x00\x00\x00", 15) == 0x034501
    
    def test_guid(self):
        raw = b"\xde\xad\xbe\xef\xfe\xed\xbe\xee\xde\xad\xbe\xef\xfe\xed\xbe\xee"
        expected = "{deadbeef-feed-be-ee-deadbeeffeedbeee}"
        assert ESEUtils.parse_ese_value(raw, 16) == expected

    def test_unsigned_short(self):
        assert ESEUtils.parse_ese_value(b"\xff\xff\xff\xff\xff\xff\xff\xff", 15) == -1
        assert ESEUtils.parse_ese_value(b"\x01\x45\x03\x00\x00\x00\x00\x00", 15) == 0x034501
    