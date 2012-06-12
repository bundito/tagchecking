import xlwt
import listreader
from datetime import date
from xl_styles import get_xf

# ICKY - temp variables... fix this!
#start_row = 0
#start_col = 0

tag_col = 0
tagdesc_col = tag_col + 1
tagmatch1_col = tag_col + 2
tagmatch2_col = tag_col + 4
descmatch1_col = tag_col + 3
descmatch2_col = tag_col + 5



# Start a new Excel sheet
wkbk = xlwt.Workbook()
sheet = wkbk.add_sheet('PID')

# Get the three sets of tags
me_tags = listreader.me_tags("ME.xls")
pid_tags = listreader.pid_tags('PID.xls')
sp3d_tags = listreader.sp3d_tags('SP3D.xls')
	
# Build a dictionary of styles from the xl_styles module
styles = ({})	
for style in ('yes', 'no', 'heading', 'tag', 'date', 'desc'):
	styles[style] = get_xf(style)


row = 0
col_widths = ({})
col_widths['tags'] = 0
col_widths['tag_desc'] = 0

# An early attempt at making this a lot less redundant

tags_main = pid_tags
tags_comp1 = sp3d_tags
tags_comp2 = me_tags
name_comp1 = 'sp3d'
name_comp2 = 'me'


col_widths[name_comp1] = 0
col_widths[name_comp2] = 0

for tag in sorted(tags_main.keys()):
		
	sheet.write(row, tag_col, tag, styles['tag'])
	sheet.write(row, tagdesc_col, tags_main[tag], styles['tag'])
	
	if len(tag) > col_widths['tags']:
		col_widths['tags'] = len(tag)
		
	if len(tags_main[tag]) > col_widths['tag_desc']:
		col_widths['tag_desc'] = len (tags_main[tag])
	
	
	
	if tag in tags_comp1:
		sheet.write(row, tagmatch1_col, 'YES', styles['yes'])
		sheet.write(row, descmatch1_col, tags_comp1[tag], styles['desc'])
		
		if len(tags_comp1[tag]) > col_widths[name_comp1]:
			col_widths[name_comp1] = len(tags_comp1[tag])
		
	else:
		sheet.write(row, tagmatch1_col, 'NO', styles['no'])
		sheet.write(row, descmatch1_col, '', styles['desc'])
		
		
	if tag in tags_comp2:
		sheet.write(row, tagmatch2_col, 'YES', styles['yes'])
		sheet.write(row, descmatch2_col, tags_comp2[tag], styles['desc'])
		
		if len(tags_comp2[tag]) > col_widths[name_comp2]:
			col_widths[name_comp2] = len(tags_comp2[tag])
		
	else:
		sheet.write(row, tagmatch2_col, 'NO', styles['no'])
		sheet.write(row, descmatch2_col, '', styles['desc'])

	row += 1

sheet.col(tag_col).width = (col_widths['tags'] + 5) * 256
sheet.col(tagdesc_col).width = (col_widths['tag_desc'] + 10) * 256
sheet.col(descmatch1_col).width = (col_widths[name_comp1] + 10) * 256
sheet.col(descmatch2_col).width = (col_widths[name_comp2] + 10) * 256

print
wkbk.save("output.xls")


#def calc_col_width(colname, tag):
#	curr_widest = col_widths[colname]
#	col_width = len(tag)
#	
#	if col_width > curr_widest:
#		cold_widths[colname] = col_width
#		
#	return col_width 
