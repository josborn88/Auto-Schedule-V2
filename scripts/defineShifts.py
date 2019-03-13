import config, openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border, Side, colors
from openpyxl.cell import Cell

class Shift:
    
	def __init__(self, name, long, restrictions, color, database, rowKnow):
		self.name = name
		self.long = long
		self.restrictions = restrictions
		self.backFill = PatternFill(start_color = color, end_color = color, fill_type = 'solid')
		self.database = database
		self.rowKnow = rowKnow

class ClosingShift:
	def __init__(self, name, color, colorTwo, database, rowKnow):
		self.name = name
		self.backFill = PatternFill(start_color = color, end_color = color, fill_type = 'solid')
		self.backFillTwo = PatternFill(start_color = colorTwo, end_color = colorTwo, fill_type = 'solid')
		self.database = database
		self.rowKnow = rowKnow
		
class ShowShift:
	def __init__(self, name, color, shiftStart, shiftEnd, database, rowKnow):
		self.name = name
		self.backFill = PatternFill(start_color = color, end_color = color, fill_type = 'solid')
		self.start = shiftStart
		self.end = shiftEnd
		self.database = database
		self.rowKnow = rowKnow