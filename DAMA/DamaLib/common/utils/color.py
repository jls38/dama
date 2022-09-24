import logging
from re import match as re_match
from enum import IntEnum

log = logging.getLogger(__name__)

class Color(IntEnum):
    BLACK = 1
    BLUE = 2
    GREEN = 3
    ORANGE = 4
    RED = 5
    WHITE = 6
    YELLOW = 7

class color_converter(object):
    def __init__(self, name:str|Color) -> None:
        self.name = name

    def hexaCode (self) -> str:
        if type(self.name) == Color:
            colorDict = self._hexaCodeDict()
            return colorDict[int(self.name)]
        elif self._isHexaCodeColor():
            return self.name[-6:]
        else:
            raise ValueError('Color not valid: expected HexaCode or Color type')

    def _hexaCodeDict (self):
        colorDict = {
        1 : '000000',
        2 : '0000ff',
        3 : '008000',
        4 : 'ffa500',
        5:  'ff0000',
        6 : 'ffffff',
        7 : 'ffff00'
        }
        return colorDict

    def _isHexaCodeColor(self):
        #Set color Hexa code format
        m = re_match("^(#?[a-fA-F0-9]{6})$", self.name) 
        
        if not m:
            return False
        
        return True