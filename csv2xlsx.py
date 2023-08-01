import os, sys, os.path, csv, tempfile, zipfile, datetime, re, base64, xml.sax.saxutils


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
	base64_string = """
		PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9InllcyI/Pg0K
		PGE6dGhlbWUgeG1sbnM6YT0iaHR0cDovL3NjaGVtYXMub3BlbnhtbGZvcm1hdHMub3JnL2RyYXdp
		bmdtbC8yMDA2L21haW4iIG5hbWU9IlRlbWEgZG8gT2ZmaWNlIj48YTp0aGVtZUVsZW1lbnRzPjxh
		OmNsclNjaGVtZSBuYW1lPSJPZmZpY2UiPjxhOmRrMT48YTpzeXNDbHIgdmFsPSJ3aW5kb3dUZXh0
		IiBsYXN0Q2xyPSIwMDAwMDAiLz48L2E6ZGsxPjxhOmx0MT48YTpzeXNDbHIgdmFsPSJ3aW5kb3ci
		IGxhc3RDbHI9IkZGRkZGRiIvPjwvYTpsdDE+PGE6ZGsyPjxhOnNyZ2JDbHIgdmFsPSI0NDU0NkEi
		Lz48L2E6ZGsyPjxhOmx0Mj48YTpzcmdiQ2xyIHZhbD0iRTdFNkU2Ii8+PC9hOmx0Mj48YTphY2Nl
		bnQxPjxhOnNyZ2JDbHIgdmFsPSI1QjlCRDUiLz48L2E6YWNjZW50MT48YTphY2NlbnQyPjxhOnNy
		Z2JDbHIgdmFsPSJFRDdEMzEiLz48L2E6YWNjZW50Mj48YTphY2NlbnQzPjxhOnNyZ2JDbHIgdmFs
		PSJBNUE1QTUiLz48L2E6YWNjZW50Mz48YTphY2NlbnQ0PjxhOnNyZ2JDbHIgdmFsPSJGRkMwMDAi
		Lz48L2E6YWNjZW50ND48YTphY2NlbnQ1PjxhOnNyZ2JDbHIgdmFsPSI0NDcyQzQiLz48L2E6YWNj
		ZW50NT48YTphY2NlbnQ2PjxhOnNyZ2JDbHIgdmFsPSI3MEFENDciLz48L2E6YWNjZW50Nj48YTpo
		bGluaz48YTpzcmdiQ2xyIHZhbD0iMDU2M0MxIi8+PC9hOmhsaW5rPjxhOmZvbEhsaW5rPjxhOnNy
		Z2JDbHIgdmFsPSI5NTRGNzIiLz48L2E6Zm9sSGxpbms+PC9hOmNsclNjaGVtZT48YTpmb250U2No
		ZW1lIG5hbWU9Ik9mZmljZSI+PGE6bWFqb3JGb250PjxhOmxhdGluIHR5cGVmYWNlPSJDYWxpYnJp
		IExpZ2h0IiBwYW5vc2U9IjAyMEYwMzAyMDIwMjA0MDMwMjA0Ii8+PGE6ZWEgdHlwZWZhY2U9IiIv
		PjxhOmNzIHR5cGVmYWNlPSIiLz48YTpmb250IHNjcmlwdD0iSnBhbiIgdHlwZWZhY2U9Iua4uOOC
		tOOCt+ODg+OCryBMaWdodCIvPjxhOmZvbnQgc2NyaXB0PSJIYW5nIiB0eXBlZmFjZT0i66eR7J2A
		IOqzoOuUlSIvPjxhOmZvbnQgc2NyaXB0PSJIYW5zIiB0eXBlZmFjZT0i562J57q/IExpZ2h0Ii8+
		PGE6Zm9udCBzY3JpcHQ9IkhhbnQiIHR5cGVmYWNlPSLmlrDntLDmmI7pq5QiLz48YTpmb250IHNj
		cmlwdD0iQXJhYiIgdHlwZWZhY2U9IlRpbWVzIE5ldyBSb21hbiIvPjxhOmZvbnQgc2NyaXB0PSJI
		ZWJyIiB0eXBlZmFjZT0iVGltZXMgTmV3IFJvbWFuIi8+PGE6Zm9udCBzY3JpcHQ9IlRoYWkiIHR5
		cGVmYWNlPSJUYWhvbWEiLz48YTpmb250IHNjcmlwdD0iRXRoaSIgdHlwZWZhY2U9Ik55YWxhIi8+
		PGE6Zm9udCBzY3JpcHQ9IkJlbmciIHR5cGVmYWNlPSJWcmluZGEiLz48YTpmb250IHNjcmlwdD0i
		R3VqciIgdHlwZWZhY2U9IlNocnV0aSIvPjxhOmZvbnQgc2NyaXB0PSJLaG1yIiB0eXBlZmFjZT0i
		TW9vbEJvcmFuIi8+PGE6Zm9udCBzY3JpcHQ9IktuZGEiIHR5cGVmYWNlPSJUdW5nYSIvPjxhOmZv
		bnQgc2NyaXB0PSJHdXJ1IiB0eXBlZmFjZT0iUmFhdmkiLz48YTpmb250IHNjcmlwdD0iQ2FucyIg
		dHlwZWZhY2U9IkV1cGhlbWlhIi8+PGE6Zm9udCBzY3JpcHQ9IkNoZXIiIHR5cGVmYWNlPSJQbGFu
		dGFnZW5ldCBDaGVyb2tlZSIvPjxhOmZvbnQgc2NyaXB0PSJZaWlpIiB0eXBlZmFjZT0iTWljcm9z
		b2Z0IFlpIEJhaXRpIi8+PGE6Zm9udCBzY3JpcHQ9IlRpYnQiIHR5cGVmYWNlPSJNaWNyb3NvZnQg
		SGltYWxheWEiLz48YTpmb250IHNjcmlwdD0iVGhhYSIgdHlwZWZhY2U9Ik1WIEJvbGkiLz48YTpm
		b250IHNjcmlwdD0iRGV2YSIgdHlwZWZhY2U9Ik1hbmdhbCIvPjxhOmZvbnQgc2NyaXB0PSJUZWx1
		IiB0eXBlZmFjZT0iR2F1dGFtaSIvPjxhOmZvbnQgc2NyaXB0PSJUYW1sIiB0eXBlZmFjZT0iTGF0
		aGEiLz48YTpmb250IHNjcmlwdD0iU3lyYyIgdHlwZWZhY2U9IkVzdHJhbmdlbG8gRWRlc3NhIi8+
		PGE6Zm9udCBzY3JpcHQ9Ik9yeWEiIHR5cGVmYWNlPSJLYWxpbmdhIi8+PGE6Zm9udCBzY3JpcHQ9
		Ik1seW0iIHR5cGVmYWNlPSJLYXJ0aWthIi8+PGE6Zm9udCBzY3JpcHQ9Ikxhb28iIHR5cGVmYWNl
		PSJEb2tDaGFtcGEiLz48YTpmb250IHNjcmlwdD0iU2luaCIgdHlwZWZhY2U9Iklza29vbGEgUG90
		YSIvPjxhOmZvbnQgc2NyaXB0PSJNb25nIiB0eXBlZmFjZT0iTW9uZ29saWFuIEJhaXRpIi8+PGE6
		Zm9udCBzY3JpcHQ9IlZpZXQiIHR5cGVmYWNlPSJUaW1lcyBOZXcgUm9tYW4iLz48YTpmb250IHNj
		cmlwdD0iVWlnaCIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBVaWdodXIiLz48YTpmb250IHNjcmlwdD0i
		R2VvciIgdHlwZWZhY2U9IlN5bGZhZW4iLz48L2E6bWFqb3JGb250PjxhOm1pbm9yRm9udD48YTps
		YXRpbiB0eXBlZmFjZT0iQ2FsaWJyaSIgcGFub3NlPSIwMjBGMDUwMjAyMDIwNDAzMDIwNCIvPjxh
		OmVhIHR5cGVmYWNlPSIiLz48YTpjcyB0eXBlZmFjZT0iIi8+PGE6Zm9udCBzY3JpcHQ9IkpwYW4i
		IHR5cGVmYWNlPSLmuLjjgrTjgrfjg4Pjgq8iLz48YTpmb250IHNjcmlwdD0iSGFuZyIgdHlwZWZh
		Y2U9IuunkeydgCDqs6DrlJUiLz48YTpmb250IHNjcmlwdD0iSGFucyIgdHlwZWZhY2U9Iuetiee6
		vyIvPjxhOmZvbnQgc2NyaXB0PSJIYW50IiB0eXBlZmFjZT0i5paw57Sw5piO6auUIi8+PGE6Zm9u
		dCBzY3JpcHQ9IkFyYWIiIHR5cGVmYWNlPSJBcmlhbCIvPjxhOmZvbnQgc2NyaXB0PSJIZWJyIiB0
		eXBlZmFjZT0iQXJpYWwiLz48YTpmb250IHNjcmlwdD0iVGhhaSIgdHlwZWZhY2U9IlRhaG9tYSIv
		PjxhOmZvbnQgc2NyaXB0PSJFdGhpIiB0eXBlZmFjZT0iTnlhbGEiLz48YTpmb250IHNjcmlwdD0i
		QmVuZyIgdHlwZWZhY2U9IlZyaW5kYSIvPjxhOmZvbnQgc2NyaXB0PSJHdWpyIiB0eXBlZmFjZT0i
		U2hydXRpIi8+PGE6Zm9udCBzY3JpcHQ9IktobXIiIHR5cGVmYWNlPSJEYXVuUGVuaCIvPjxhOmZv
		bnQgc2NyaXB0PSJLbmRhIiB0eXBlZmFjZT0iVHVuZ2EiLz48YTpmb250IHNjcmlwdD0iR3VydSIg
		dHlwZWZhY2U9IlJhYXZpIi8+PGE6Zm9udCBzY3JpcHQ9IkNhbnMiIHR5cGVmYWNlPSJFdXBoZW1p
		YSIvPjxhOmZvbnQgc2NyaXB0PSJDaGVyIiB0eXBlZmFjZT0iUGxhbnRhZ2VuZXQgQ2hlcm9rZWUi
		Lz48YTpmb250IHNjcmlwdD0iWWlpaSIgdHlwZWZhY2U9Ik1pY3Jvc29mdCBZaSBCYWl0aSIvPjxh
		OmZvbnQgc2NyaXB0PSJUaWJ0IiB0eXBlZmFjZT0iTWljcm9zb2Z0IEhpbWFsYXlhIi8+PGE6Zm9u
		dCBzY3JpcHQ9IlRoYWEiIHR5cGVmYWNlPSJNViBCb2xpIi8+PGE6Zm9udCBzY3JpcHQ9IkRldmEi
		IHR5cGVmYWNlPSJNYW5nYWwiLz48YTpmb250IHNjcmlwdD0iVGVsdSIgdHlwZWZhY2U9IkdhdXRh
		bWkiLz48YTpmb250IHNjcmlwdD0iVGFtbCIgdHlwZWZhY2U9IkxhdGhhIi8+PGE6Zm9udCBzY3Jp
		cHQ9IlN5cmMiIHR5cGVmYWNlPSJFc3RyYW5nZWxvIEVkZXNzYSIvPjxhOmZvbnQgc2NyaXB0PSJP
		cnlhIiB0eXBlZmFjZT0iS2FsaW5nYSIvPjxhOmZvbnQgc2NyaXB0PSJNbHltIiB0eXBlZmFjZT0i
		S2FydGlrYSIvPjxhOmZvbnQgc2NyaXB0PSJMYW9vIiB0eXBlZmFjZT0iRG9rQ2hhbXBhIi8+PGE6
		Zm9udCBzY3JpcHQ9IlNpbmgiIHR5cGVmYWNlPSJJc2tvb2xhIFBvdGEiLz48YTpmb250IHNjcmlw
		dD0iTW9uZyIgdHlwZWZhY2U9Ik1vbmdvbGlhbiBCYWl0aSIvPjxhOmZvbnQgc2NyaXB0PSJWaWV0
		IiB0eXBlZmFjZT0iQXJpYWwiLz48YTpmb250IHNjcmlwdD0iVWlnaCIgdHlwZWZhY2U9Ik1pY3Jv
		c29mdCBVaWdodXIiLz48YTpmb250IHNjcmlwdD0iR2VvciIgdHlwZWZhY2U9IlN5bGZhZW4iLz48
		L2E6bWlub3JGb250PjwvYTpmb250U2NoZW1lPjxhOmZtdFNjaGVtZSBuYW1lPSJPZmZpY2UiPjxh
		OmZpbGxTdHlsZUxzdD48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48L2E6
		c29saWRGaWxsPjxhOmdyYWRGaWxsIHJvdFdpdGhTaGFwZT0iMSI+PGE6Z3NMc3Q+PGE6Z3MgcG9z
		PSIwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6bHVtTW9kIHZhbD0iMTEwMDAwIi8+PGE6
		c2F0TW9kIHZhbD0iMTA1MDAwIi8+PGE6dGludCB2YWw9IjY3MDAwIi8+PC9hOnNjaGVtZUNscj48
		L2E6Z3M+PGE6Z3MgcG9zPSI1MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOmx1bU1v
		ZCB2YWw9IjEwNTAwMCIvPjxhOnNhdE1vZCB2YWw9IjEwMzAwMCIvPjxhOnRpbnQgdmFsPSI3MzAw
		MCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIg
		dmFsPSJwaENsciI+PGE6bHVtTW9kIHZhbD0iMTA1MDAwIi8+PGE6c2F0TW9kIHZhbD0iMTA5MDAw
		Ii8+PGE6dGludCB2YWw9IjgxMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0Pjxh
		OmxpbiBhbmc9IjU0MDAwMDAiIHNjYWxlZD0iMCIvPjwvYTpncmFkRmlsbD48YTpncmFkRmlsbCBy
		b3RXaXRoU2hhcGU9IjEiPjxhOmdzTHN0PjxhOmdzIHBvcz0iMCI+PGE6c2NoZW1lQ2xyIHZhbD0i
		cGhDbHIiPjxhOnNhdE1vZCB2YWw9IjEwMzAwMCIvPjxhOmx1bU1vZCB2YWw9IjEwMjAwMCIvPjxh
		OnRpbnQgdmFsPSI5NDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9hOmdzPjxhOmdzIHBvcz0iNTAwMDAi
		PjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48YTpzYXRNb2QgdmFsPSIxMTAwMDAiLz48YTpsdW1N
		b2QgdmFsPSIxMDAwMDAiLz48YTpzaGFkZSB2YWw9IjEwMDAwMCIvPjwvYTpzY2hlbWVDbHI+PC9h
		OmdzPjxhOmdzIHBvcz0iMTAwMDAwIj48YTpzY2hlbWVDbHIgdmFsPSJwaENsciI+PGE6bHVtTW9k
		IHZhbD0iOTkwMDAiLz48YTpzYXRNb2QgdmFsPSIxMjAwMDAiLz48YTpzaGFkZSB2YWw9Ijc4MDAw
		Ii8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PC9hOmdzTHN0PjxhOmxpbiBhbmc9IjU0MDAwMDAiIHNj
		YWxlZD0iMCIvPjwvYTpncmFkRmlsbD48L2E6ZmlsbFN0eWxlTHN0PjxhOmxuU3R5bGVMc3Q+PGE6
		bG4gdz0iNjM1MCIgY2FwPSJmbGF0IiBjbXBkPSJzbmciIGFsZ249ImN0ciI+PGE6c29saWRGaWxs
		PjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIi8+PC9hOnNvbGlkRmlsbD48YTpwcnN0RGFzaCB2YWw9
		InNvbGlkIi8+PGE6bWl0ZXIgbGltPSI4MDAwMDAiLz48L2E6bG4+PGE6bG4gdz0iMTI3MDAiIGNh
		cD0iZmxhdCIgY21wZD0ic25nIiBhbGduPSJjdHIiPjxhOnNvbGlkRmlsbD48YTpzY2hlbWVDbHIg
		dmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6cHJzdERhc2ggdmFsPSJzb2xpZCIvPjxhOm1p
		dGVyIGxpbT0iODAwMDAwIi8+PC9hOmxuPjxhOmxuIHc9IjE5MDUwIiBjYXA9ImZsYXQiIGNtcGQ9
		InNuZyIgYWxnbj0iY3RyIj48YTpzb2xpZEZpbGw+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiLz48
		L2E6c29saWRGaWxsPjxhOnByc3REYXNoIHZhbD0ic29saWQiLz48YTptaXRlciBsaW09IjgwMDAw
		MCIvPjwvYTpsbj48L2E6bG5TdHlsZUxzdD48YTplZmZlY3RTdHlsZUxzdD48YTplZmZlY3RTdHls
		ZT48YTplZmZlY3RMc3QvPjwvYTplZmZlY3RTdHlsZT48YTplZmZlY3RTdHlsZT48YTplZmZlY3RM
		c3QvPjwvYTplZmZlY3RTdHlsZT48YTplZmZlY3RTdHlsZT48YTplZmZlY3RMc3Q+PGE6b3V0ZXJT
		aGR3IGJsdXJSYWQ9IjU3MTUwIiBkaXN0PSIxOTA1MCIgZGlyPSI1NDAwMDAwIiBhbGduPSJjdHIi
		IHJvdFdpdGhTaGFwZT0iMCI+PGE6c3JnYkNsciB2YWw9IjAwMDAwMCI+PGE6YWxwaGEgdmFsPSI2
		MzAwMCIvPjwvYTpzcmdiQ2xyPjwvYTpvdXRlclNoZHc+PC9hOmVmZmVjdExzdD48L2E6ZWZmZWN0
		U3R5bGU+PC9hOmVmZmVjdFN0eWxlTHN0PjxhOmJnRmlsbFN0eWxlTHN0PjxhOnNvbGlkRmlsbD48
		YTpzY2hlbWVDbHIgdmFsPSJwaENsciIvPjwvYTpzb2xpZEZpbGw+PGE6c29saWRGaWxsPjxhOnNj
		aGVtZUNsciB2YWw9InBoQ2xyIj48YTp0aW50IHZhbD0iOTUwMDAiLz48YTpzYXRNb2QgdmFsPSIx
		NzAwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpzb2xpZEZpbGw+PGE6Z3JhZEZpbGwgcm90V2l0aFNo
		YXBlPSIxIj48YTpnc0xzdD48YTpncyBwb3M9IjAiPjxhOnNjaGVtZUNsciB2YWw9InBoQ2xyIj48
		YTp0aW50IHZhbD0iOTMwMDAiLz48YTpzYXRNb2QgdmFsPSIxNTAwMDAiLz48YTpzaGFkZSB2YWw9
		Ijk4MDAwIi8+PGE6bHVtTW9kIHZhbD0iMTAyMDAwIi8+PC9hOnNjaGVtZUNscj48L2E6Z3M+PGE6
		Z3MgcG9zPSI1MDAwMCI+PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnRpbnQgdmFsPSI5ODAw
		MCIvPjxhOnNhdE1vZCB2YWw9IjEzMDAwMCIvPjxhOnNoYWRlIHZhbD0iOTAwMDAiLz48YTpsdW1N
		b2QgdmFsPSIxMDMwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48YTpncyBwb3M9IjEwMDAwMCI+
		PGE6c2NoZW1lQ2xyIHZhbD0icGhDbHIiPjxhOnNoYWRlIHZhbD0iNjMwMDAiLz48YTpzYXRNb2Qg
		dmFsPSIxMjAwMDAiLz48L2E6c2NoZW1lQ2xyPjwvYTpncz48L2E6Z3NMc3Q+PGE6bGluIGFuZz0i
		NTQwMDAwMCIgc2NhbGVkPSIwIi8+PC9hOmdyYWRGaWxsPjwvYTpiZ0ZpbGxTdHlsZUxzdD48L2E6
		Zm10U2NoZW1lPjwvYTp0aGVtZUVsZW1lbnRzPjxhOm9iamVjdERlZmF1bHRzLz48YTpleHRyYUNs
		clNjaGVtZUxzdC8+PGE6ZXh0THN0PjxhOmV4dCB1cmk9InswNUE0QzI1Qy0wODVFLTQzNDAtODVB
		My1BNTUzMUU1MTBEQjJ9Ij48dGhtMTU6dGhlbWVGYW1pbHkgeG1sbnM6dGhtMTU9Imh0dHA6Ly9z
		Y2hlbWFzLm1pY3Jvc29mdC5jb20vb2ZmaWNlL3RoZW1lbWwvMjAxMi9tYWluIiBuYW1lPSJPZmZp
		Y2UgVGhlbWUiIGlkPSJ7NjJGOTM5QjYtOTNBRi00REI4LTlDNkItRDZDN0RGREM1ODlGfSIgdmlk
		PSJ7NEEzQzQ2RTgtNjFDQy00NjAzLUE1ODktNzQyMkE0N0E4RTRBfSIvPjwvYTpleHQ+PC9hOmV4
		dExzdD48L2E6dGhlbWU+
	"""
	oneline_base64_string = base64_string.translate(str.maketrans('', '', '\t\n\r'))
	utf8_bytes = base64.b64decode(oneline_base64_string)
	result_string = utf8_bytes.decode('utf-8')
	return result_string


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
		result_string += '<si><t>{}</t></si>'.format(xml.sax.saxutils.escape(csv_string))
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
	if os.path.isfile(xlsx_file_name):
		raise Exception("file '{}' already exists".format(xlsx_file_name))
	temp_file_name = ''
	try:
		workbook_strings = {}
		sheets = []
		for csv_file_name in csv_file_names:
			print("reading '{}'".format(csv_file_name))
			sheet_headers, sheet_widths, sheet_rows = read_csv(csv_file_name, delimiter, workbook_strings)
			sheets.append((len(sheets) + 1, os.path.splitext(os.path.basename(csv_file_name))[0][:30], sheet_headers, sheet_widths, sheet_rows))
		with tempfile.NamedTemporaryFile(delete=False) as temp_file_handle:
			temp_file_name = temp_file_handle.name
			print("creating '{}'".format(temp_file_name))
			with zipfile.ZipFile(temp_file_handle, mode='w') as zip_file_handle:
				print("writing header files")
				zip_file_handle.writestr('[Content_Types].xml', gen_content_types_xml())
				zip_file_handle.writestr('_rels/.rels', gen__rels_dotrels())
				zip_file_handle.writestr('docProps/app.xml', gen_docProps_app_xml(sheets))
				zip_file_handle.writestr('docProps/core.xml', gen_docProps_core_xml(author_name, datetime.datetime.now()))
				zip_file_handle.writestr('xl/_rels/workbook.xml.rels', gen_xl_rels_workbook_xml_rels(sheets))
				zip_file_handle.writestr('xl/theme/theme1.xml', gen_xl_theme_theme_xml())
				zip_file_handle.writestr('xl/styles.xml', gen_xl_styles_style_xml())
				zip_file_handle.writestr('xl/workbook.xml', gen_xl_workbook_xml(sheets))
				print('writing shared strings')
				zip_file_handle.writestr('xl/sharedStrings.xml', gen_xl_shared_strings_xml(workbook_strings))
				for sheet in sheets:
					sheet_id, sheet_name, sheet_headers, sheet_widths, sheet_rows = sheet
					print("writing sheet '{}'".format(sheet_name))
					zip_file_handle.writestr('xl/worksheets/sheet{}.xml'.format(sheet_id), gen_xl_worksheets_sheetN_xml(sheet_headers, sheet_widths, sheet_rows, sheet_id == 1))
		print("moving '{}' to '{}'".format(temp_file_name, xlsx_file_name))
		with open(temp_file_name, 'rb') as src_file, open(xlsx_file_name, 'wb') as dst_file:
			dst_file.write(src_file.read())
	finally:
		if os.path.isfile(temp_file_name):
			print("removing '{}'".format(temp_file_name))
			os.remove(temp_file_name)
	print('done')

# parameters
xlsx_file_name = sys.argv[1] + '.xlsx'
csv_file_names = sys.argv[2:]
delimiter = ','
author_name = 'author_name'

# entry point
if __name__ == "__main__":
	main()

