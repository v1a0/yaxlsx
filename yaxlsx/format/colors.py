class FormatCellsColor(str):
    """
    Colors

    Palette based on https://en.wikipedia.org/wiki/Web_colors
    """

    black = '#000000'
    white = '#FFFFFF'

    red = '#FF0000'
    green = '#008000'
    blue = '#0000FF'

    yellow = '#FFFF00'
    purple = '#A020F0'
    grey = '#808080'

    def __init__(self, hex_value: str = None, red: int = 0, green: int = 0, blue: int = 0):
        if hex_value is None:
            hex_value = self.rgb_to_hex(red=red, green=green, blue=blue)

        super().__init__(hex_value)

    @staticmethod
    def rgb_to_hex(red: int, green: int, blue: int):
        if (red > 255 or red < 0) or (green > 255 or green < 0) or (blue > 255 or blue < 0):
            raise ValueError(f"Invalid color code ({red},{green},{blue}): "
                             f"Max value for color (225,255,255), min value for color (0,0,0),")

        return f"#{hex(red)[2:]:0>2}{hex(green)[2:]:0>2}{hex(blue)[2:]:0>2}".upper()
