import xlwt
import gettags
from datetime import date

# Get the three sets of tags
pid_tags = gettags.pid_tags("../WD Lists/PID.xls")
sp3d_tags = gettags.sp3d_tags("../WD Lists/SP3D.xls")
me_tags = gettags.me_tags2("../WD Lists/ME.xls")

# Start a new Excel sheet
wkbk = xlwt.Workbook()
sheet = wkbk.add_sheet('Sheet1')

# Set up formats to make the finished report pretty.

# XF's are Excel "extended formatting" objects. I wish I'd found
# the 'easyxf' approach a day or so earlier... these are actually
# easy to read and change!
heading_template = """
borders: left thin, top thin, right thin, bottom thick;
align: horizontal center;
font: bold True;
"""
heading_xf = xlwt.easyxf(heading_template)

tag_template = """
borders: left thin, top thin, right thin, bottom thin;
align: horizontal center;
"""
tag_xf = xlwt.easyxf(tag_template)

yes_template = """
borders: left thin, top thin, right thin, bottom thin;
align: horizontal center;
pattern: pattern fine-dots, fore_color green;
""" 
yes_xf = xlwt.easyxf(yes_template)

no_template = """
borders: left thin, top thin, right thin, bottom thin;
align: horizontal center;
pattern: pattern fine-dots, fore_color red;
"""
no_xf = xlwt.easyxf(no_template)

title_template = """
font: bold True, height 350;
align: horizontal center;
"""
title_xf = xlwt.easyxf(title_template)

date_template = """
align: horizontal center;
"""
date_xf = xlwt.easyxf(date_template)


# A whole bunch of paramters that need to be in a config file instead.
start_row = 4
title_column = 4
start_col = 0
sp3d_tag_column = 4
me_tag_column = 8

today = date.today()
today = today.strftime('%m/%d/%y')

# Excel column widths are measured in 1/256th's the width of the '0' character.
# If that's not weird enough, the width property seems to expect three values(?)
# This needs to be investigated further (dig into module source)
sheet.col(start_col).width = 1600 + 50 * 50
sheet.col(sp3d_tag_column).width = 1600 + 50 * 150
sheet.col(me_tag_column).width = 1600 + 50 * 50

# Writing to the worksheets is easy - (row, col, data, format)

sheet.write(0, title_column, 'Equipment Comparison Report', title_xf)
sheet.write(1, title_column - 1, 'Unnamed Power Plant', date_xf)
sheet.write(1, title_column + 1, 'Report created: ' + today, date_xf)

# A bunch of column headings
sheet.write(start_row - 1, start_col, 'PID Tags', heading_xf)
sheet.write(start_row - 1, start_col + 1, 'In SP3D?', heading_xf)
sheet.write(start_row - 1, start_col + 2, 'In Mech?', heading_xf)

sheet.write(start_row - 1, sp3d_tag_column, 'SP3D Tags', heading_xf)
sheet.write(start_row - 1, sp3d_tag_column + 1, 'In P&ID?', heading_xf)
sheet.write(start_row - 1, sp3d_tag_column + 2, 'In Mech?', heading_xf)

sheet.write(start_row - 1, me_tag_column, 'Mech Engr Tags', heading_xf)
sheet.write(start_row - 1, me_tag_column + 1, 'In P&ID?', heading_xf)
sheet.write(start_row - 1, me_tag_column + 2, 'In SP3D?', heading_xf)


# This needs rework. Too redundant. All it's doing is taking one list of 
# tags at a time and comparing each entry against the other two lists.

for tag in pid_tags:
	sheet.write((pid_tags.index(tag) + start_row), start_col, tag, tag_xf)
	
	if sp3d_tags.count(tag) != 0:
		sheet.write(pid_tags.index(tag) + start_row, start_col + 1, 'YES', yes_xf)
	else:
		sheet.write(pid_tags.index(tag) + start_row, start_col + 1, 'NO', no_xf)
		
	if me_tags.count(tag) != 0:
		sheet.write(pid_tags.index(tag) + start_row, start_col + 2, 'YES', yes_xf)
	else:
		sheet.write(pid_tags.index(tag) + start_row, start_col + 2, 'NO', no_xf)

for tag in sp3d_tags:
	sheet.write((sp3d_tags.index(tag) + start_row), sp3d_tag_column, tag, tag_xf)
	
	if pid_tags.count(tag) != 0:
		sheet.write(sp3d_tags.index(tag) + start_row, sp3d_tag_column + 1, 'YES', yes_xf)
	else:
		sheet.write(sp3d_tags.index(tag) + start_row, sp3d_tag_column + 1, 'NO', no_xf)
		
	if me_tags.count(tag) != 0:
		sheet.write(sp3d_tags.index(tag) + start_row, sp3d_tag_column + 2, 'YES', yes_xf)
	else:
		sheet.write(sp3d_tags.index(tag) + start_row, sp3d_tag_column + 2, 'NO', no_xf)
	
for tag in me_tags:
	sheet.write((me_tags.index(tag) + start_row), me_tag_column, tag, tag_xf)
	
	if pid_tags.count(tag) != 0:
		sheet.write(me_tags.index(tag) + start_row, me_tag_column + 1, 'YES', yes_xf)
	else:
		sheet.write(me_tags.index(tag) + start_row, me_tag_column + 1, 'NO', no_xf)

	if sp3d_tags.count(tag) != 0:
		sheet.write(me_tags.index(tag) + start_row, me_tag_column + 2, 'YES', yes_xf)
	else:
		sheet.write(me_tags.index(tag) + start_row, me_tag_column + 2, 'NO', no_xf)
	

# Fin. Close the Excel file.	
wkbk.save("output.xls")