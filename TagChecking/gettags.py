#
# gettags.py - Read item tags from the three differet equipment reports
#
# Last revision: 24 May 2011 by SPH
#
# Load up the super-handy Excel-reading module (xlrd)
# And the perl coder's all-time favorite, regular expressions (re)
import xlrd
import re

import pdb

def pid_tags(filename, tag_column = 1, start_row = 6):

	# Open the desired workbook, select "Sheet1" and get the values in TAG_COLUMN
	wkbk = xlrd.open_workbook(filename)
	sheet = wkbk.sheet_by_name('Sheet1')
	tags = sheet.col_values(colx=tag_column, start_rowx=start_row, end_rowx = None)

	# Sort ASCII-ascending; also moves any empty cells to the front of the list
	tags.sort()

	# How many empty cells are there?
	empty_count = tags.count('')

	# Delete 'em from the list, starting from position 0 through the number found
	del tags[0:empty_count]

	# Start a new list to hold the reformatted tags
	new_tags = []

	# Loop through and reformat all the tags
	for tag in tags:
		# Split up the old-style tag so we can rearrange the individual parts
		tag_parts = re.split('-', tag)	
		
		# Rebuild the tag in the "new & improved" style
		new_tag = '{2}-{1}-{0}-{3}'.format(*tag_parts)

		# Add to the list, assuming it doesn't already exist (???)
		if new_tags.count(new_tag) == 0:
			new_tags.append(new_tag)

	# Send the sorted and reformatted tags back to the main program	
	new_tags.sort()
	return new_tags

#---------#

# Get tags from the SP3D report. They shouldn't need any reformatting if the
# piping designers did their part correctly...

def sp3d_tags(filename, tag_column = 1, start_row = 8):
	
	# Open the desired workbook, select "Sheet1" and get the values in TAG_COLUMN
	wkbk = xlrd.open_workbook(filename)
	sheet = wkbk.sheet_by_name('Sheet1')
	tags = sheet.col_values(colx=tag_column, start_rowx=start_row, end_rowx = None)

	# Sort ASCII-ascending; also moves any empty cells to the front of the list
	tags.sort()

	# How many empty cells are there?
	empty_count = tags.count('')

	# Delete 'em from the list, starting from position 0 through the number found
	del tags[0:empty_count]
	
	# As before, make sure there aren't any duplicates...
	# Add to the list, assuming it doesn't already exist (???)
	
	new_tags=[]
	
	for tag in tags:
		if new_tags.count(tag) == 0:
			new_tags.append(tag)
			
			
	return new_tags
	
#--------#

# Get tags from Mechanical Engineering's list. They never do anything the simple way,
# so this is the most complex of the three. 

def me_tags(filename, start_col = 1, start_row = 2, wkbk_sheet_index = 1):
	
	# We need to turn on formatting, because the engineers like to leave old things
	# in their list, and just apply 'strikeout' to the text.
	wkbk = xlrd.open_workbook(filename, formatting_info = True)
	sheet = wkbk.sheet_by_index(wkbk_sheet_index)

	tags = set([])

	# They also don't list their tags in a single column, and xlrd lacks the
	# ability to select a 2D rectangle of Excel cells. Hence, a double loop.
	for col in range(1,sheet.ncols):
		for row in range(2,sheet.nrows):
			this_cell = sheet.cell(row, col)
			type_of_cell = sheet.cell_type(row, col)
			
			# Check for blank/empty cells and skip 'em
			# 0 = XL_CELL_EMPTY, 6 = XL_CELL_BLANK (unclear on the diff)
			if type_of_cell not in [0, 6]:
								
				# Sub-check - is the text struck out? The engineers like to
				# leave old things in their list and just apply "strikeout"
				struck = is_struck(wkbk, sheet, row, col)
				
				if struck is False:
					tags.add(this_cell.value)
					
	
	
	# Now that we have a (partially) clean set of tags, it's time to look
	# for more of the engineers' quirks. 
	
	addable_cells = set([])
	removable_cells = set([])
	
	# STEP 1: more than one entry per cell - we'll look for newlines
	# and split them up.
	# (All the multi-line cells have only had 2 entries. This needs to be
	#  modified in case they start adding more. It's on the to-do list.)
	
	for tag in tags:
		
		match = re.split('\n', tag)
		
		if match is not None and len(match) != 1:
			# We check to be sure there's more than one match because
			# sometimes there's a newline left at the end with no more text
			removable_cells.add(tag)
			addable_cells.add(match[0])
			addable_cells.add(match[1])
			
	# Add the multiline cells to the set of tags
	tags = set(tags).union(addable_cells)
	# And then cut out the original multiline cells
	tags = set(tags).difference(removable_cells)
	
	
	# STEP 2: The engineers list redundant equipment in one cell, adding
	# a suffix like "A/B/C" instead of putting A on one row, B on the next...
	# There should be a space after the numbering ends, so that's our break
	addable_cells = set([])
	removable_cells = set([])
	
	for tag in tags:
	
		
	
		parts = re.match('(.*0\d) (.*)', tag)
	
		if parts is not None:
			
			main_tag = parts.group(1)
			subtags = re.split('/', parts.group(2))
			
			
			
			for subtag in subtags:
				addable_cells.add(main_tag + subtag)
			
			removable_cells.add(tag)
	
	# Repeat the add/subtract routine on the main set of tags...
	
	# pdb.set_trace()
	
	union_tags = set(tags).union(addable_cells)
	
	# pdb.set_trace()
	
	tags = union_tags
	
	diff_tags = set(tags).difference(removable_cells)
	
	tags = diff_tags
	
	# pdb.set_trace()
	
	# Needed sets for the booleans - convert back to list like other def's
	tag_list = list(tags)
	tag_list.sort()
	
	# And that should do it... we're done cleaning up their tags!
	return tag_list
	
def me_tags2(filename, tag_column = 1, start_row = 6, wkbk_sheet_index = 0):
	# We need to turn on formatting, because the engineers like to leave old things
	# in their list, and just apply 'strikeout' to the text.
	wkbk = xlrd.open_workbook(filename, formatting_info = True)
	sheet = wkbk.sheet_by_index(wkbk_sheet_index)
	
	tags = set([])
	
	column_contents = sheet.col_values(colx=tag_column, start_rowx=start_row, end_rowx = None)
	
	for row in range(start_row, len(column_contents)):
		this_cell = sheet.cell(row, tag_column)
		type_of_cell = sheet.cell_type(row, tag_column)
		
	# Check for blank/empty cells and skip 'em
		# 0 = XL_CELL_EMPTY, 6 = XL_CELL_BLANK (unclear on the diff)
		if type_of_cell not in [0, 6]:
							
			# Sub-check - is the text struck out? The engineers like to
			# leave old things in their list and just apply "strikeout"
			struck = is_struck(wkbk, sheet, row, tag_column)
			
			if struck is False:
				tags.add(this_cell.value)
	
	
	# Now that we have a (partially) clean set of tags, it's time to look
	# for more of the engineers' quirks. 
	
	addable_cells = set([])
	removable_cells = set([])
	
	# STEP 1: more than one entry per cell - we'll look for newlines
	# and split them up.
	# (All the multi-line cells have only had 2 entries. This needs to be
	#  modified in case they start adding more. It's on the to-do list.)
	
	for tag in tags:
		
		match = re.split('\n', tag)
		
		if match is not None and len(match) != 1:
			# We check to be sure there's more than one match because
			# sometimes there's a newline left at the end with no more text
			removable_cells.add(tag)
			addable_cells.add(match[0])
			addable_cells.add(match[1])
			
	# Add the multiline cells to the set of tags
	tags = set(tags).union(addable_cells)
	# And then cut out the original multiline cells
	tags = set(tags).difference(removable_cells)
	
	
	# STEP 2: The engineers list redundant equipment in one cell, adding
	# a suffix like "A/B/C" instead of putting A on one row, B on the next...
	# There should be a space after the numbering ends, so that's our break
	addable_cells = set([])
	removable_cells = set([])
	
	for tag in tags:
	
		
	
		parts = re.match('(.*0\d) (.*)', tag)
	
		if parts is not None:
			
			main_tag = parts.group(1)
			subtags = re.split('/', parts.group(2))
			
			
			
			for subtag in subtags:
				addable_cells.add(main_tag + subtag)
			
			removable_cells.add(tag)
	
	# Repeat the add/subtract routine on the main set of tags...
	
	# pdb.set_trace()
	
	union_tags = set(tags).union(addable_cells)
	
	# pdb.set_trace()
	
	tags = union_tags
	
	diff_tags = set(tags).difference(removable_cells)
	
	tags = diff_tags
	
	# pdb.set_trace()
	
	# Needed sets for the booleans - convert back to list like other def's
	tag_list = list(tags)
	tag_list.sort()
	
	# And that should do it... we're done cleaning up their tags!
	return tag_list
	
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