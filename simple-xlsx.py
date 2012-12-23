import sys
import re
import os
from datetime import datetime
import shutil
import zipfile
from xml.sax.saxutils import escape

###Functions
#passing this function the column count will return the corresponding
#column lettering, as it appears in Excel.  
def _getColumnLetter(column_number):
	letters = {
		-1 : '',
		0 : 'A',
		1 : 'B',
		2 : 'C',
		3 : 'D',
		4 : 'E',
		5 : 'F',
		6 : 'G',
		7 : 'H',
		8 : 'I',
		9 : 'J',
		10 : 'K',
		11 : 'L',
		12 : 'M',
		13 : 'N',
		14 : 'O',
		15 : 'P',
		16 : 'Q',
		17 : 'R',
		18 : 'S',
		19 : 'T',
		20 : 'U',
		21 : 'V',
		22 : 'W',
		23 : 'X',
		24 : 'Y',
		25 : 'Z'
	}
	if column_number < 703:
		num1 = -1
	else:
		num1 = ((column_number-27)//676)-1
	num2 = (((column_number-1)//26)-1) - ((num1 + 1) * 26) 
	num3 = (column_number-1)%26
	column = letters[num1] + letters[num2] + letters[num3]
	return column

#This function will be used to generate a list and dictionary of the unique strings in the workbook
#to be used in creating the sharedStrings.xml file.
def _addSharedString(shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, element):
	tag = ' ' #used for shared strings, empty by default
	try:
		float(element)
	except ValueError:
		total_shared_strings += 1
		if element in shared_strings_dictionary:
			element = shared_strings_dictionary[element]
		else:
			shared_strings_dictionary[element] = unique_shared_strings #add it to the dictionary...
			shared_strings_list.append(element) #...and the list
			element = unique_shared_strings
			unique_shared_strings += 1 #iterate after adding, first value must be 0
		tag = ' t="s"' #shared strings require this attribute
	return shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, element, tag


#Gets the width of each worksheet column, measured by the character length of the largest
#cell in the column.  This is used to generate a starting worksheet view where the visible 
#column width fits the data.
def _getColumnWidths(column_widths, elements, reset=False):
	if reset == True:
		column_widths = []
	else:
		for i in range(len(elements)):
			try:
				if len(elements[i]) > column_widths[i]:
					column_widths[i] = len(elements[i])
				else:
					pass
			except IndexError:
				column_widths.append(len(elements[i]))
	return column_widths



#This function will replace non-Ascii characters with question marks
def _fixNonAscii(s):
	x = ""
	for i in s:
		if ord(i)<128:
			x+=i
		else:
			x+='?'
	return x


#This function will create the XML for a given worksheet cell.  This includes special
#formatting and shared strings.
def _createCellData(shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, elements, column):
	value = escape(_fixNonAscii(elements[column]))
	element_to_write = ''
	t_tag = ''
	(shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, element_to_write, t_tag) = _addSharedString(shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, value) #create shared string
	return shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, elements, element_to_write, t_tag


def writeWorkbook(input_files, output_file, delimiter_name='tab'):
	
	###Check the arguments
	#Delimiter
	delimiter_options = {
		'tab': '\t',
		'comma': ',',
		'colon': ':',
		'semicolon': ';'
	}
	try:
		delimiter = delimiter_options[delimiter_name]
	except KeyError:
		raise Warning('Invalid delimiter.  Acceptable delimiters are:\n%s'%(delimiter_options.keys()))
	
	#Output file
	'''THIS NEEDS TO ACCOUNT FOR WINDOWS AND LINUX, AS WELL AS ARGUMENTS WITH NO/PARTIAL PATHS'''
	if '/' in output_file:
		(working_directory, output_file_name) = output_file.rsplit('/', 1)
	else:
		(working_directory, output_file_name) = output_file.rsplit('\\', 1)
	
	if not os.path.exists(working_directory):
		raise Warning('Invalid working directory: %s does not exist!'%(working_directory))
		
	output_file_name = re.sub('\.\w+$', '', output_file_name)
	
	#Input files
	for f in input_files:
		if not os.path.exists(f):
			raise Warning('Invalid input file: %s does not exist!'%(f))
	
	###Variable Declaration
	username = ''
	date_created = datetime.isoformat(datetime.now()) #Used for 'Date created' attribute
	shared_strings_dictionary = {} #used for sharedStrings.xml 
	shared_strings_list = [] #used for sharedStrings.xml 
	written_sheets = [] #list of the names written worksheets, as they will be displayed
	column_widths = [] #for setting column width in the worksheets
	worksheet_number = 0 #used to keep track of the number of written worksheets, as opposed to the number of user-supplied worksheet filenames
	total_shared_strings = 0 #used for sharedStrings.xml 
	unique_shared_strings = 0 #used for sharedStrings.xml 

	###Files and Folders
	#Create a temp folder to build the worbook in
	new_workbook_directory = working_directory + r'\workbookTemp' #temporary dir for workbook xml files
	try:
		shutil.rmtree(new_workbook_directory) #Remove the temp directory, if it exists
	except:
		pass
	os.mkdir(new_workbook_directory)

	try:
		###Create all of the standard files first.  These are always the same.
		###Create [Content_Types].xml
		content_types_filename = new_workbook_directory + r'\[Content_Types].xml'
		with open(content_types_filename, mode='w') as content_types:
			content_types.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')

		###Create \_rels\.rels
		rels_filename = new_workbook_directory + r'\.rels'
		with open(rels_filename, mode='w') as rels:
			rels.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')

		###Create styles.xml file
		styles_filename = new_workbook_directory + r'\styles.xml'
		with open(styles_filename, mode='w') as styles:
			styles.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="19"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="18"/><color theme="3"/><name val="Cambria"/><family val="2"/><scheme val="major"/></font><font><b/><sz val="15"/><color theme="3"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="13"/><color theme="3"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="3"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF9C6500"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FF3F3F76"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color rgb="FF3F3F3F"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color rgb="FFFA7D00"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FFFA7D00"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><i/><sz val="11"/><color rgb="FF7F7F7F"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><sz val="11"/><color theme="0"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><u /><sz val="11" /><color theme="10" /><name val="Calibri" /><family val="2" /></font></fonts><fills count="33"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFCC99"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFA5A5A5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFCC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="5" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="6" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="7" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="8" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.79998168889431442"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.59999389629810485"/><bgColor indexed="65"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.39997558519241921"/><bgColor indexed="65"/></patternFill></fill></fills><borders count="10"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color theme="4"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="thick"><color theme="4" tint="0.499984740745262"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="medium"><color theme="4" tint="0.39997558519241921"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FF7F7F7F"/></left><right style="thin"><color rgb="FF7F7F7F"/></right><top style="thin"><color rgb="FF7F7F7F"/></top><bottom style="thin"><color rgb="FF7F7F7F"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FF3F3F3F"/></left><right style="thin"><color rgb="FF3F3F3F"/></right><top style="thin"><color rgb="FF3F3F3F"/></top><bottom style="thin"><color rgb="FF3F3F3F"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="double"><color rgb="FFFF8001"/></bottom><diagonal/></border><border><left style="double"><color rgb="FF3F3F3F"/></left><right style="double"><color rgb="FF3F3F3F"/></right><top style="double"><color rgb="FF3F3F3F"/></top><bottom style="double"><color rgb="FF3F3F3F"/></bottom><diagonal/></border><border><left style="thin"><color rgb="FFB2B2B2"/></left><right style="thin"><color rgb="FFB2B2B2"/></right><top style="thin"><color rgb="FFB2B2B2"/></top><bottom style="thin"><color rgb="FFB2B2B2"/></bottom><diagonal/></border><border><left/><right/><top style="thin"><color theme="4"/></top><bottom style="double"><color theme="4"/></bottom><diagonal/></border></borders><cellStyleXfs count="43"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="3" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="4" fillId="0" borderId="2" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="5" fillId="0" borderId="3" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="5" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="6" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="7" fillId="3" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="8" fillId="4" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="9" fillId="5" borderId="4" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="10" fillId="6" borderId="5" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="11" fillId="6" borderId="4" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="12" fillId="0" borderId="6" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="13" fillId="7" borderId="7" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="14" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="8" borderId="8" applyNumberFormat="0" applyFont="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="15" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="16" fillId="0" borderId="9" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="9" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="10" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="11" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="12" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="13" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="14" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="15" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="16" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="17" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="18" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="19" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="20" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="21" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="22" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="23" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="24" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="25" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="26" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="27" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="28" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="29" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="30" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="1" fillId="31" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="17" fillId="32" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/><xf numFmtId="0" fontId="18" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"><alignment vertical="top" /><protection locked="0" /></xf></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="17" fillId="9" borderId="0" xfId="18" /><xf numFmtId="0" fontId="18" fillId="0" borderId="0" xfId="42" applyAlignment="1" applyProtection="1" /></cellXfs><cellStyles count="43"><cellStyle name="20% - Accent1" xfId="19" builtinId="30" customBuiltin="1"/><cellStyle name="20% - Accent2" xfId="23" builtinId="34" customBuiltin="1"/><cellStyle name="20% - Accent3" xfId="27" builtinId="38" customBuiltin="1"/><cellStyle name="20% - Accent4" xfId="31" builtinId="42" customBuiltin="1"/><cellStyle name="20% - Accent5" xfId="35" builtinId="46" customBuiltin="1"/><cellStyle name="20% - Accent6" xfId="39" builtinId="50" customBuiltin="1"/><cellStyle name="40% - Accent1" xfId="20" builtinId="31" customBuiltin="1"/><cellStyle name="40% - Accent2" xfId="24" builtinId="35" customBuiltin="1"/><cellStyle name="40% - Accent3" xfId="28" builtinId="39" customBuiltin="1"/><cellStyle name="40% - Accent4" xfId="32" builtinId="43" customBuiltin="1"/><cellStyle name="40% - Accent5" xfId="36" builtinId="47" customBuiltin="1"/><cellStyle name="40% - Accent6" xfId="40" builtinId="51" customBuiltin="1"/><cellStyle name="60% - Accent1" xfId="21" builtinId="32" customBuiltin="1"/><cellStyle name="60% - Accent2" xfId="25" builtinId="36" customBuiltin="1"/><cellStyle name="60% - Accent3" xfId="29" builtinId="40" customBuiltin="1"/><cellStyle name="60% - Accent4" xfId="33" builtinId="44" customBuiltin="1"/><cellStyle name="60% - Accent5" xfId="37" builtinId="48" customBuiltin="1"/><cellStyle name="60% - Accent6" xfId="41" builtinId="52" customBuiltin="1"/><cellStyle name="Accent1" xfId="18" builtinId="29" customBuiltin="1"/><cellStyle name="Accent2" xfId="22" builtinId="33" customBuiltin="1"/><cellStyle name="Accent3" xfId="26" builtinId="37" customBuiltin="1"/><cellStyle name="Accent4" xfId="30" builtinId="41" customBuiltin="1"/><cellStyle name="Accent5" xfId="34" builtinId="45" customBuiltin="1"/><cellStyle name="Accent6" xfId="38" builtinId="49" customBuiltin="1"/><cellStyle name="Bad" xfId="7" builtinId="27" customBuiltin="1"/><cellStyle name="Calculation" xfId="11" builtinId="22" customBuiltin="1"/><cellStyle name="Check Cell" xfId="13" builtinId="23" customBuiltin="1"/><cellStyle name="Explanatory Text" xfId="16" builtinId="53" customBuiltin="1"/><cellStyle name="Good" xfId="6" builtinId="26" customBuiltin="1"/><cellStyle name="Heading 1" xfId="2" builtinId="16" customBuiltin="1"/><cellStyle name="Heading 2" xfId="3" builtinId="17" customBuiltin="1"/><cellStyle name="Heading 3" xfId="4" builtinId="18" customBuiltin="1"/><cellStyle name="Heading 4" xfId="5" builtinId="19" customBuiltin="1"/><cellStyle name="Hyperlink" xfId="42" builtinId="8" /><cellStyle name="Input" xfId="9" builtinId="20" customBuiltin="1"/><cellStyle name="Linked Cell" xfId="12" builtinId="24" customBuiltin="1"/><cellStyle name="Neutral" xfId="8" builtinId="28" customBuiltin="1"/><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Note" xfId="15" builtinId="10" customBuiltin="1"/><cellStyle name="Output" xfId="10" builtinId="21" customBuiltin="1"/><cellStyle name="Title" xfId="1" builtinId="15" customBuiltin="1"/><cellStyle name="Total" xfId="17" builtinId="25" customBuiltin="1"/><cellStyle name="Warning Text" xfId="14" builtinId="11" customBuiltin="1"/></cellStyles><dxfs count="1"><dxf><font><condense val="0" /><extend val="0" /><color rgb="FF9C0006" /></font><fill><patternFill><bgColor rgb="FFFFC7CE" /></patternFill></fill></dxf></dxfs><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>')

		###Create xl\.theme1.xml file
		theme_filename = new_workbook_directory + r'\theme1.xml'
		with open(theme_filename, mode='w') as theme:
			theme.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="MS P????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')


		###Now we create the custom files.
		###Write Worksheets
		for input_filename in input_files:
			with open(input_filename) as input_file:
				worksheet_number += 1
				worksheet_row_count = 0
				column_names = [] #for use in conditional formatting
				
				column_widths = _getColumnWidths(column_widths, '', True) #reset the column widths array
				
				#count the number of rows in the file and the number of columns to get the worksheet range
				for line in input_file:
					worksheet_row_count += 1
					elements = re.split(delimiter, _fixNonAscii(line))
					if worksheet_row_count == 1:
						worksheet_range = _getColumnLetter(len(elements)) 
						column_names = elements
					column_widths = _getColumnWidths(column_widths, elements)
				worksheet_range += str(worksheet_row_count)
				input_file.seek(0)
				
				#Open the output sheet.xml file for writing
				worksheet_filename = new_workbook_directory + r'\sheet' + str(worksheet_number) + '.xml'
				with open(worksheet_filename, mode='w') as final_worksheet:
					worksheet_row_count = 0 #start the count over
					final_worksheet.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n<dimension ref="A1:%s" />\n<sheetViews>\n<sheetView workbookViewId="0">\n</sheetView>\n</sheetViews>\n<sheetFormatPr defaultRowHeight="15" />\n<cols>\n'%(worksheet_range))
					for column in range(len(column_widths)): #Add the column widths
						final_worksheet.write(u'<col min="%d" max="%d" width="%f" bestFit="1" customWidth="1" />\n'%(column+1, column+1, column_widths[column]*1.2))
					final_worksheet.write(u'</cols>\n<sheetData>\n') 
					for line in input_file: #iterate through each line of the file
						worksheet_row_count += 1
						elements = re.split('\t', line) #Split the line by the tabs
						number_of_columns = len(elements)
						final_worksheet.write(u'<row r="%d" spans="1:%d">\n'%(worksheet_row_count, number_of_columns))
						for j in range(number_of_columns): #For each element of each line...
							cell_number = _getColumnLetter(j+1) + str(worksheet_row_count)
							(shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, elements, element_to_write, tags_to_write) = _createCellData(shared_strings_dictionary, total_shared_strings, unique_shared_strings, shared_strings_list, elements, j)
							final_worksheet.write(u'<c r="%s"%s>\n<v>%s</v>\n</c>\n'%(cell_number, tags_to_write, element_to_write))
						final_worksheet.write(u'</row>\n')
					final_worksheet.write(u'</sheetData>\n') 
					final_worksheet.write(u'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />\n</worksheet>')
				
				#Add the worksheet name to the list
				match = re.search('[\\\/](\w+)\.\w+$', input_filename) #get the file name without extension...
				written_sheets.append(match.group(1)) 

		###Create sharedStrings.xml file
		shared_strings_filename = new_workbook_directory + r'\sharedStrings.xml'
		with open(shared_strings_filename, mode='w') as shared_strings:
			unique_shared_strings = len(shared_strings_list)
			shared_strings.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">\n'%(total_shared_strings, unique_shared_strings))
			for item in shared_strings_list:
				if item != '':
					shared_strings.write(u'<si>\n<t>%s</t>\n</si>\n'%(item))
				else:
					shared_strings.write(u'<si>\n<t/>\n</si>\n')
			shared_strings.write(u'</sst>')


		###Create app.xml file
		app_xml_filename = new_workbook_directory + r'\app.xml'
		with open(app_xml_filename, mode='w') as app_xml:
			app_xml.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">\n  <DocSecurity>0</DocSecurity>\n<ScaleCrop>false</ScaleCrop>\n<HeadingPairs>\n<vt:vector size="2" baseType="variant">\n<vt:variant>\n<vt:lpstr>Worksheets</vt:lpstr>\n</vt:variant>\n<vt:variant>\n<vt:i4>%d</vt:i4>\n</vt:variant>\n</vt:vector>\n</HeadingPairs>\n<TitlesOfParts>\n<vt:vector size="%d" baseType="lpstr">\n'%(worksheet_number, worksheet_number))
			for i in range(len(written_sheets)):
				app_xml.write(u'<vt:lpstr>%s</vt:lpstr>\n'%(written_sheets[i]))
			app_xml.write(u'</vt:vector>\n</TitlesOfParts>\n<LinksUpToDate>false</LinksUpToDate>\n<SharedDoc>false</SharedDoc>\n<HyperlinksChanged>false</HyperlinksChanged>\n<AppVersion>12.0000</AppVersion>\n</Properties>')


		###Create core.xml file
		core_xml_filename = new_workbook_directory + r'\core.xml'
		with open(core_xml_filename, mode='w') as core_xml:
			core_xml.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n<dc:creator>Will\'s Python XLSX Writer</dc:creator>\n<cp:lastModifiedBy>%s</cp:lastModifiedBy>\n<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>\n<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>\n</cp:coreProperties>'%(username, date_created, date_created))


		###Create workbook.xml.rels file
		wxr_filename = new_workbook_directory + r'\workbook.xml.rels'
		with open(wxr_filename, mode='w') as wxr:
			wxr.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n')
			for i in range(worksheet_number):
				wxr.write(u'<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml" />\n'%(i+1, i+1))
			wxr.write(u'<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />\n<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />\n<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />\n</Relationships>'%(worksheet_number+3, worksheet_number+2, worksheet_number+1))

		###Create workbook.xml
		workbook_xml_filename = new_workbook_directory + r'\workbook.xml'
		with open(workbook_xml_filename, mode='w') as workbook_xml:
			workbook_xml.write(u'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4506" />\n<workbookPr defaultThemeVersion="124226" />\n<bookViews>\n<workbookView xWindow="0" yWindow="120" windowWidth="18075" windowHeight="6660" activeTab="0" />\n</bookViews>\n<sheets>\n')
			for i in range(len(written_sheets)):
				workbook_xml.write(u'<sheet name="%s" sheetId="%d" r:id="rId%d" />\n'%(written_sheets[i], i+1, i+1))
			workbook_xml.write(u'</sheets>\n<calcPr calcId="0" />\n</workbook>')

		###Now we create the archive for the XLSX file
		new_excel_file = os.path.join(working_directory, output_file_name+'.xlsx')
		try:
			os.remove(new_excel_file) #Erase it if one already exists
		except:
			pass
		zip = zipfile.ZipFile(new_excel_file, "w", zipfile.ZIP_DEFLATED) #create a new zipfile object
		zip.write(content_types_filename, r'[Content_Types].xml')
		zip.write(rels_filename, r'_rels\.rels')
		zip.write(styles_filename, r'xl\styles.xml')
		zip.write(theme_filename, r'xl\theme\theme1.xml')
		zip.write(shared_strings_filename, r'xl\sharedStrings.xml')
		zip.write(app_xml_filename, r'docProps\app.xml')
		zip.write(core_xml_filename, r'docProps\core.xml')
		zip.write(wxr_filename, r'xl\_rels\workbook.xml.rels')
		zip.write(workbook_xml_filename, r'xl\workbook.xml')

		for i in range(worksheet_number):
			temp_worksheet_dir = new_workbook_directory + r'\sheet' + str(i+1) + '.xml'
			archive_worksheet_dir = r'xl\worksheets\sheet' + str(i+1) + '.xml'
			zip.write(temp_worksheet_dir, archive_worksheet_dir)
		zip.close()

	###Erase temp files and directories
	finally:
		shutil.rmtree(new_workbook_directory) #Remove the temp directory




if __name__ == '__main__':
	from optparse import OptionParser

	#Handle command line arguments
	usage = "Generates and Excel .xlsx workbook file using tab-delimited text input file.  Each input file becomes its own worksheet, named using the input file name.\n\nusage: %prog [options] output_filename input_file1 input_file2 etc ..."
	parser = OptionParser(usage = usage)
	parser.add_option("-d", "--delimiter", dest="delimiter", default="tab", help="Delimiter used to separate values in the input files.  This must be consistent accross all input files.  Options include 'tab', 'comma', 'colon' and 'semicolon'.  The default is 'tab'")
	(options, args) = parser.parse_args()
	
	#Generate the workbook
	if len(args) < 2:
		print usage
	else:
		writeWorkbook(args[1:], args[0], options.delimiter)
