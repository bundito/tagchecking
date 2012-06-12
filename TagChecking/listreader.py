import xlrd
import re


def me_tags(filename, tag_column=1, start_row=6, wkbk_sheet_index=0):
	# We need to turn on formatting, because the engineers like to leave old things
	# in their list, and just apply 'strikeout' to the text.
	wkbk = xlrd.open_workbook(filename, formatting_info=True)
	sheet = wkbk.sheet_by_index(wkbk_sheet_index)
	
	tags = set([])
	items = dict({})
	
	column_contents = sheet.col_values(colx=tag_column, start_rowx=start_row, end_rowx=None)
	
	for row in range(start_row, len(column_contents)):
		this_cell = sheet.cell(row, tag_column)
		type_of_cell = sheet.cell_type(row, tag_column)
		
	# Check for blank/empty cells and skip 'em
		# 0 = XL_CELL_EMPTY, 6 = XL_CELL_BLANK (unclear on the diff)
		if type_of_cell not in [0, 6]:
							
			# Sub-check - is the text struck out? The engineers like to
			# leave old things in their list and just apply "strikeout"
			struck = is_struck(wkbk, sheet, row, tag_column)
			
			# Look for all-alpha labels instead of standard alpha-numeric tag
			all_alpha = re.match('^\D*$', this_cell.value)

#			print this_cell.value, " x ", all_alpha

			if struck is False and bool(all_alpha) is False:
				tags.add(this_cell.value)
				description = sheet.cell(row, tag_column + 3).value
				items[this_cell.value] = description
	
	
	# Now that we have a (partially) clean set of tags, it's time to look
	# for more of the engineers' quirks. 
	
	# STEP 1: more than one entry per cell - we'll look for newlines
	# and split them up.
	# (All the multi-line cells have only had 2 entries. This needs to be
	#  modified in case they start adding more. It's on the to-do list.)
	
	for tag in tags:
		
		match = re.split('\n', tag)
		
		if match is not None and len(match) != 1:
			# We check to be sure there's more than one match because
			# sometimes there's a newline left at the end with no more text
			match0 = match[0]
			match1 = match[1]
			desc = items[tag]

			items[match0] = desc
			items[match1] = desc
			del items[tag]			

	# STEP 2: The engineers list redundant equipment in one cell, adding
	# a suffix like "A/B/C" instead of putting A on one row, B on the next...
	# There should be a space after the numbering ends, so that's our break
	
	# Make a copy of our dictionary, since we can't change it while iter'ing
	new_items = items.copy()
	
	for tag in iter(items):
		
		desc = new_items[tag]
#		desc = desc.decode()
#		desc = desc.uppercase()
		new_items[tag] = desc.upper()
		
		# Multitags SHOULD be "...01 A" - a digit followed by a space
		# They may have done more weird stuff... must keep checking
		parts = re.match('(.*0\d) (.*)', tag)
		
		if parts is not None:
		
			main_tag = parts.group(1)
			subtags = re.split('/', parts.group(2))
			desc = items[tag]
			
			for subtag in subtags:
				
				single_tag = main_tag + subtag
				new_items[single_tag] = desc.upper()
			
			del new_items[tag]


	return new_items 

######################################

def pid_tags(filename, tag_col = 1, desc_col = 7, start_row = 6):
	
	# Open the desired workbook, select "Sheet1" and get the values in TAG_COLUMN
	wkbk = xlrd.open_workbook(filename)
	sheet = wkbk.sheet_by_name('Sheet1')
	col_contents = sheet.col_values(colx=tag_col, start_rowx=start_row, end_rowx = None)
	
	items = ({})
	
	for row in range(start_row, start_row + len(col_contents)):
		this_cell = sheet.cell(row, tag_col)
		type_of_cell = sheet.cell_type(row, tag_col)
		
		if type_of_cell == 1:
			tag = reformat_pid_tag(this_cell.value)
			desc = sheet.cell(row, desc_col).value
			desc = re.sub('(\n|\r)', ' ', desc)
			items[tag] = desc
		
	return items

##################################

def sp3d_tags(filename, tag_col = 1, desc_col = 2, start_row = 8):
	
	# Open the desired workbook, select "Sheet1" and get the values in TAG_COLUMN
	wkbk = xlrd.open_workbook(filename)
	sheet = wkbk.sheet_by_name('Sheet1')
	col_contents = sheet.col_values(colx=tag_col, start_rowx=start_row, end_rowx = None)
	
	items = ({})
	
	for row in range(start_row, start_row + len(col_contents)):
		tag = sheet.cell(row, tag_col).value
		desc = sheet.cell(row, desc_col).value
		type_of_cell = sheet.cell_type(row, tag_col)
		
		if type_of_cell == 1:
			items[tag] = desc
			
	return items
	
	
######################################
# Utility functions
######################################
	
	
def is_struck(wkbk, sheet, row, col):
	# Extracting the formatting data to see if the text has "strikeout" applied
	# involves multiple steps and objects. It bounces between numerous objects,
	# so the object names are commented in all caps to avoid confusion.
	# (XF stands for "Extended Formatting", in case you're curious.)

	# xf_index = numeric index to the WORKSHEETS's list of format (XF) objects
	# Each cell has one, so we grab it from the individual cell
	this_xf_index = sheet.cell_xf_index(row, col)
		
	# xf_list = retrieve the specific format (XF) object from the WORKBOOK's list
	this_xf = wkbk.xf_list[this_xf_index]
	
	# font_index = the FORMAT OBJECT's numeric index that refers to the 
	# WORKBOOK's list of font objects
	this_font_index = this_xf.font_index
	
	# font_list = retrieve the specific FONT OBJECT from the WORKBOOK's list
	# We finally get to a FONT OBJECT that contains various format properties
	this_font_obj = wkbk.font_list[this_font_index]
				
	# One of those FONT OBJECT properties is struck_out (0/1) for true/false
	# I prefer True and False for clarity, so we convert it
	strikeout = bool(this_font_obj.struck_out)

	return strikeout





def reformat_pid_tag (tag):
	# Moved this to it's own function in the event it changes, or the PID
	# report eventually gets reworked and this will be unnecessary.
	
	# Split up the old-style tag so we can rearrange the individual parts
	tag_parts = re.split('-', tag)	
	
	if len(tag_parts) != 4:
		return tag
	
	# Rebuild the tag in the "new & improved" style
	new_tag = '{2}-{1}-{0}-{3}'.format(*tag_parts)
	
	return new_tag