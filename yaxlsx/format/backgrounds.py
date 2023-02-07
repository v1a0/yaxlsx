from .colors import FormatCellsColor


class FormatCellsBackground(str):
    __color__ = FormatCellsColor

    black = {'bg_color': __color__.black}
    white = {'bg_color': __color__.white}

    red = {'bg_color': __color__.red}
    green = {'bg_color': __color__.green}
    blue = {'bg_color': __color__.blue}
