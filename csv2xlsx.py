import os, sys, os.path, csv, tempfile, zipfile, datetime, re


def excel_column_name(column_index):
	letters = []
	column_index += 1
	while True:
		column_index, remainder = divmod(column_index, 26)
		letters.insert(0, chr(64 + remainder))
		if column_index == 0:
			break
	return ''.join(letters)


def remove_field_suffix(field_name):
	if re.match('^.*_[gsx]$', field_name):
		return field_name[:-2]
	else:
		return field_name


def string_is_number(string):
	try:
		float(string)
		return True
	except ValueError:
		return False


def read_csv(file_name, delimiter, workbook_strings):
	sheet_headers = []
	sheet_formats = {}
	sheet_withts = {}
	sheet_rows = []
	workbook_string_index = len(workbook_strings)
	with open(file_name, 'r') as file_handle:
		csv_reader = csv.DictReader(file_handle, delimiter=delimiter)
		for csv_field_name in csv_reader.fieldnames:
			if not csv_field_name.endswith('_x'):
				sheet_header_name = remove_field_suffix(csv_field_name)
				if workbook_strings.get(sheet_header_name) is None:
					workbook_strings[sheet_header_name] = workbook_string_index
					sheet_header_index = workbook_string_index
					workbook_string_index += 1
				else:
					sheet_header_index = workbook_strings[sheet_header_name]
				if csv_field_name.endswith('_g'):
					sheet_formats[sheet_header_index] = 'g'
				else:
					sheet_formats[sheet_header_index] = 's'
				len_sheet_header_name = len(sheet_header_name) + 2
				if sheet_withts.get(sheet_header_index, 0) < len_sheet_header_name:
					sheet_withts[sheet_header_index] = len_sheet_header_name * 1.25
				sheet_headers.append(sheet_header_index)
		for csv_row in csv_reader:
			sheet_row = {}
			for csv_field_name in csv_row:
				sheet_header_name = remove_field_suffix(csv_field_name)
				sheet_header_index = workbook_strings.get(sheet_header_name)
				if sheet_header_index is not None:
					cell_value = csv_row[csv_field_name]
					if (sheet_formats[sheet_header_index] == 'g') and string_is_number(cell_value):
						sheet_row[sheet_header_index] = str(cell_value)
					else:
						if workbook_strings.get(cell_value) is None:
							workbook_strings[cell_value] = workbook_string_index
							sheet_row[sheet_header_index] = int(workbook_string_index)
							workbook_string_index += 1
						else:
							sheet_row[sheet_header_index] = int(workbook_strings[cell_value])
					len_cell_value = len(cell_value)
					if sheet_withts.get(sheet_header_index, 0) < len_cell_value:
						sheet_withts[sheet_header_index] = len_cell_value * 1.25
			sheet_rows.append(sheet_row)
	return sheet_headers, sheet_withts, sheet_rows


def gen_content_types_xml():
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
			<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
			<Default Extension="xml" ContentType="application/xml" />
			<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
			<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
			<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
			<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
			<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
			<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
			<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
		</Types>
	"""
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen__rels_dotrels():
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
		</Relationships>
	"""
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_docProps_app_xml(sheets):
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
			<Application>Microsoft Excel</Application>
			<DocSecurity>0</DocSecurity>
			<ScaleCrop>false</ScaleCrop>
			<HeadingPairs>
				<vt:vector size="2" baseType="variant">
					<vt:variant>
						<vt:lpstr>Sheets</vt:lpstr>
					</vt:variant>
					<vt:variant>
						<vt:i4>{}</vt:i4>
					</vt:variant>
				</vt:vector>
			</HeadingPairs>
			<TitlesOfParts>
				<vt:vector size="{}" baseType="lpstr">
	""".format(len(sheets), len(sheets))
	for sheet in sheets:
		_, sheet_name, _, _, _ = sheet
		result_string += '<vt:lpstr>{}</vt:lpstr>'.format(sheet_name)
	result_string += """
				</vt:vector>
			</TitlesOfParts>
			<LinksUpToDate>false</LinksUpToDate>
			<SharedDoc>false</SharedDoc>
			<HyperlinksChanged>false</HyperlinksChanged>
			<AppVersion>16.0300</AppVersion>
		</Properties>
	"""
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_docProps_core_xml(author_name, date_time):
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
			<dc:creator>{}</dc:creator>
			<cp:lastModifiedBy>{}</cp:lastModifiedBy>
			<dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>
			<dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>
		</cp:coreProperties>
	""".format(author_name, author_name, date_time.isoformat(), date_time.isoformat())
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_xl_rels_workbook_xml_rels(sheets):
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	"""
	for sheet in sheets:
		sheet_id, _, _, _, _ = sheet
		result_string += '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml" />'.format(sheet_id, sheet_id)
	result_string += '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />'.format(len(sheets) + 1)
	result_string += '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />'.format(len(sheets) + 2)
	result_string += '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />'.format(len(sheets) + 3)
	result_string += '</Relationships>'
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_xl_theme_theme_xml():
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
		</a:theme>
	"""
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_xl_styles_style_xml():
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
			<fonts count="2" x14ac:knownFonts="1">
				<font>
					<sz val="11" />
					<color theme="1" />
					<name val="Calibri" />
					<family val="2" />
					<scheme val="minor" />
				</font>
				<font>
					<b />
					<sz val="11" />
					<color theme="1" />
					<name val="Calibri" />
					<family val="2" />
					<scheme val="minor" />
				</font>
			</fonts>
			<fills count="1">
				<fill>
					<patternFill patternType="none" />
				</fill>
			</fills>
			<borders count="1">
				<border>
					<left />
					<right />
					<top />
					<bottom />
					<diagonal />
				</border>
			</borders>
			<cellStyleXfs count="2">
				<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
				<xf numFmtId="0" fontId="1" fillId="0" borderId="0" />
			</cellStyleXfs>
			<cellXfs count="2">
				<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
				<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" />
			</cellXfs>
			<cellStyles count="2">
				<cellStyle name="Normal" xfId="0" builtinId="0" />
				<cellStyle name="Normal 2" xfId="1" />
			</cellStyles>
			<dxfs count="0" />
		</styleSheet>
	"""
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_xl_workbook_xml(sheets):
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
			<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420" />
			<workbookPr filterPrivacy="1" defaultThemeVersion="164011" />
	"""
	result_string += '<sheets>'
	for sheet in sheets:
		sheet_id, sheet_name, _, _, _ = sheet
		result_string += '<sheet name="{}" sheetId="{}" r:id="rId{}" />'.format(sheet_name, sheet_id, sheet_id)
	result_string += '</sheets>'
	result_string += r"""
			<extLst>
				<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
					<x15:workbookPr chartTrackingRefBase="1" />
				</ext>
			</extLst>
		</workbook>
	"""
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_xl_shared_strings_xml(workbook_strings):
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">
	""".format(len(workbook_strings), len(workbook_strings))
	for csv_string in sorted(workbook_strings, key=workbook_strings.get):
		result_string += '<si><t>{}</t></si>'.format(csv_string)
	result_string += '</sst>'
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def gen_xl_worksheets_sheetN_xml(sheet_headers, sheet_widths, sheet_rows, tab_selected):
	result_string = """
		<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
		<dimension ref="A1:{}{}" />
		<sheetViews>
			<sheetView {} workbookViewId="0">
				<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" />
				<selection pane="bottomLeft" />
			</sheetView>
		</sheetViews>
		<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25" />
		<cols>
	""".format(excel_column_name(len(sheet_headers) - 1), len(sheet_rows) + 1, 'tabSelected="1"' if tab_selected else '')
	for sheet_header_index, sheet_header in enumerate(sheet_headers):
		result_string += '<col min="{}" max="{}" width="{}" customWidth="1" />'.format(sheet_header_index + 1, sheet_header_index + 1, sheet_widths[sheet_header])
	result_string += """
		</cols>
		<sheetData>
	"""
	row_string = '<row r="1" spans="1:{}" x14ac:dyDescent="0.25">'.format(len(sheet_headers))
	for sheet_header_index, sheet_header in enumerate(sheet_headers):
		row_string += '<c r="{}1" s="1" t="s"><v>{}</v></c>'.format(excel_column_name(sheet_header_index), sheet_header)
	row_string += '</row>'
	result_string += row_string
	for sheet_row_index, sheet_row in enumerate(sheet_rows):
		row_string = '<row r="{}" spans="1:{}" x14ac:dyDescent="0.25">'.format(sheet_row_index + 2, len(sheet_headers))
		for column_index, sheet_header in enumerate(sheet_row):
			cell_value = sheet_row[sheet_header]
			if isinstance(cell_value, int):
				row_string += '<c r="{}{}" t="s"><v>{}</v></c>'.format(excel_column_name(column_index), sheet_row_index + 2, cell_value)
			else:
				row_string += '<c r="{}{}"><v>{}</v></c>'.format(excel_column_name(column_index), sheet_row_index + 2, cell_value)
		row_string += '</row>'
		result_string += row_string
	result_string += """
		</sheetData>
		<autoFilter ref="A1:{}1" />
		</worksheet>
	""".format(excel_column_name(len(sheet_headers) - 1))
	return result_string.translate(str.maketrans('', '', '\t\n\r'))


def main():
	temp_file_name = ''
	if os.path.isfile(xlsx_file_name):
		raise Exception("file '{}' already exists".format(xlsx_file_name))
	try:
		workbook_strings = {}
		sheets = []
		for csv_file_name in csv_file_names:
			sheet_headers, sheet_widths, sheet_rows = read_csv(csv_file_name, delimiter, workbook_strings)
			sheets.append((len(sheets) + 1, os.path.splitext(os.path.basename(csv_file_name))[0][:30], sheet_headers, sheet_widths, sheet_rows))
		with tempfile.NamedTemporaryFile(delete=False) as temp_file_handle:
			temp_file_name = temp_file_handle.name
			with zipfile.ZipFile(temp_file_handle, mode='w') as zip_file_handle:
				zip_file_handle.writestr('[Content_Types].xml', gen_content_types_xml())
				zip_file_handle.writestr('_rels/.rels', gen__rels_dotrels())
				zip_file_handle.writestr('docProps/app.xml', gen_docProps_app_xml(sheets))
				zip_file_handle.writestr('docProps/core.xml', gen_docProps_core_xml(author_name, datetime.datetime.now()))
				zip_file_handle.writestr('xl/_rels/workbook.xml.rels', gen_xl_rels_workbook_xml_rels(sheets))
				zip_file_handle.writestr('xl/theme/theme1.xml', gen_xl_theme_theme_xml())
				zip_file_handle.writestr('xl/styles.xml', gen_xl_styles_style_xml())
				zip_file_handle.writestr('xl/workbook.xml', gen_xl_workbook_xml(sheets))
				zip_file_handle.writestr('xl/sharedStrings.xml', gen_xl_shared_strings_xml(workbook_strings))
				for sheet in sheets:
					sheet_id, _, sheet_headers, sheet_widths, sheet_rows = sheet
					zip_file_handle.writestr('xl/worksheets/sheet{}.xml'.format(sheet_id), gen_xl_worksheets_sheetN_xml(sheet_headers, sheet_widths, sheet_rows, sheet_id == 1))
		with open(temp_file_name, 'rb') as src_file, open(xlsx_file_name, 'wb') as dst_file:
			dst_file.write(src_file.read())
	finally:
		if os.path.isfile(temp_file_name):
			os.remove(temp_file_name)


# parameters
xlsx_file_name = sys.argv[1] + '.xlsx'
csv_file_names = sys.argv[2:]
delimiter = ','
author_name = 'author_name'

# entry point
if __name__ == "__main__":
	main()
