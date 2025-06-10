import pandas as pd
import xlrd
import re

def preprocess(sheet):
	"""
	Preprocess the xls data to segregate table names and data values.
	"""
	all_text = list()
	# rows, cols = sheet.nrows, sheet.ncols
	rows = sheet.nrows
	for row_idx in range(rows):
		cell_value = sheet.row_values(row_idx)
		all_text.append(cell_value)  # extract complete text content linearly from the file
	tables = list()
	len_text = len(all_text)
	data_vals = list()
	final_data = list()
	indices = list()
	for i in range(2, len_text):  # The first two rows contain titles/description of the data so can be skipped 
		if(i == len_text-1):
			final_data.append(data_vals)

		if(any([isinstance(x, (int, float)) for x in all_text[i]])):
			# print("numbers here")
			temp_list = all_text[i]
			data_vals.append(temp_list)
		elif(all([isinstance(x, (str)) for x in all_text[i]])):
			final_data.append(data_vals)
			data_vals = list()
			# final_format = dict.fromkeys([x for x in all_text[i] if x != '' ])
			indices = [t for t, x in enumerate(all_text[i]) if x != '']
			indices.append(len(all_text[i]))
			tables.append([x for x in all_text[i]])
			# print(tables, indices)
			# data_vals = 
	return tables, final_data[1:]

def convert_final(tables:list, final_data:list) -> dict:
	"""
	Format the data into one structure where keys and values represent tables and their respective data.
	"""
	final_format = dict()
	flag = False
	for table, data in zip(tables, final_data):
		# extract/segregate data indices from horizontally positioned tables
		if(sum([1 for x in table if x != '']) > 1):  
			flag = True
			indices = [t for t, x in enumerate(table) if x != '']
			indices.append(len(table))
			print(indices)
		else:
			flag = False

		# extract/segregate table name and data from horizontally positioned tables
		if(flag):
			obj = dict()
			# print(table, data)
			obj = {table[ind]:list() for ind in indices[:-1]}
			for i, dat in enumerate(data):
				data_list = list()
				for ind in range(len(indices) - 1):
					pair = (data[i][indices[ind]], data[i][indices[ind]+2])
					# print(pair)
					obj[table[indices[ind]]].append(pair)
			# print(table[ind],obj)
			final_format.update(obj)

		else:
			obj = dict()
			title = [x for x in table if x != '']
			if(len(title) == 0):
				continue
			obj = {title[0]:list()}
			# print(obj, data)
			for i, dat in enumerate(data):
				numbers = [item for item in dat if isinstance(item, (int, float))]
				strings = [item for item in dat if isinstance(item, str) and item != '']
				if(len(strings) == 0):
					continue
				pair = (strings[0], numbers)
				obj[title[0]].append(pair)
			final_format.update(obj)

	final_format = {k: v for k, v in final_format.items() if v}
	return final_format


def load_excel_sheets(path: str) -> dict:
	"""
	Load all sheets in Excel as a dictionary of DataFrames.
	"""
	try:
		wb = xlrd.open_workbook('Data/capbudg.xls')
		sheet = wb.sheet_by_index(0)
		tables, final_data = preprocess(sheet=sheet)
		final_format = convert_final(tables=tables, final_data=final_data)
		# print(final_format)
		return final_format
	except Exception as e:
		raise RuntimeError(f"Failed to load Excel file: {e}")

# load_excel_sheets("Data\capbudg.xls")


def get_row_names(table):
	"""
	Return the first column values as row names.
	"""
	all_tables = [x for x,y in table]
	return all_tables


def calculate_row_sum(table, row_name: str):
	"""
	Return sum of numeric values in the specified row.
	"""
	data = dict(table)
	row_sum = sum(data[row_name])
	# print(row_sum)
	return f"{row_sum:.2f}"
