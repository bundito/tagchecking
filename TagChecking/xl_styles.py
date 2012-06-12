import xlwt

# Set up formats to make the finished report pretty.

def get_xf(style):
	
	# XF's are Excel "extended formatting" objects. I wish I'd found
	# the 'easyxf' approach a day or so earlier... these are actually
	# easy to read and change!
	
	if style == "heading":
		template = """
			borders: left thin, top thin, right thin, bottom thick;
			align: horizontal center;
			font: bold True;
		"""
	
	elif style == "tag":
		template = """
			borders: left thin, top thin, right thin, bottom thin;
			align: horizontal center;
		"""
	
	elif style == "yes":
		template = """
			borders: left thin, top thin, right thin, bottom thin;
			align: horizontal center;
			pattern: pattern fine-dots, fore_color green;
		""" 
	
	elif style == "no":
		template = """
			borders: left thin, top thin, right thin, bottom thin;
			align: horizontal center;
			pattern: pattern fine-dots, fore_color red;
		"""
	
	elif style == "title":
		template = """
			font: bold True, height 350;
			align: horizontal center; 
		"""


	elif style == "desc":
		template = """
			align: horizontal center;
			borders: left thin, top thin, right thin, bottom thin;
		"""
		
	else:
		template = """
		"""
	
	
	return xlwt.easyxf(template)
