
# The purpose of this file is to house all cell fill colors that will be used when restoring each sheet's advanced formatting.

from openpyxl.styles import PatternFill


# Function for filling an Excel cell with a specified color. It takes the color as input. 
def cell_fill(color_choice):
    return PatternFill(start_color=color_choice, end_color=color_choice, fill_type='solid')


# Each color is defined by enacting the cell_fill() function with the specified color's hex code as a parameter.
blue = cell_fill('6495ED')
magenta = cell_fill('FF028D')
cyan = cell_fill('00FFFF')
plum = cell_fill('DDA0DD')
yellow = cell_fill('FFFF00')
ut_orange = cell_fill('FDB44E')
brown = cell_fill('644117')
yellow_green = cell_fill('9ACD32')
lime_green = cell_fill('80CD32')
red = cell_fill('FF0000')
black = cell_fill('000000')
grey = cell_fill('808080')
silver = cell_fill('F8F8FF')
green = cell_fill('00FF00')
light_green = cell_fill('90EE90')
light_red = cell_fill('F08080')
coral = cell_fill('FF8383')
white = cell_fill('FFFFFF')
