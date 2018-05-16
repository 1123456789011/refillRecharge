import re
from openpyxl import load_workbook
"""
This code automates recharges using the openpyxl library. The script takes data
from a compiled utility data excel sheet and inserts them into the recharge excel
sheets. The script takes the data from specific locations, so the data and the
recharge sheets need to stay in a consistent format, or data will be input improperly.
See dataCompilation.xlsx for formating example.

Usage:
	Gets data from a data compilation excel sheet to store into Utility Recharges,
	Dining Utility Recharges, and Housing Recharges excel sheets. To change which
	to input and output, go to individual methods and change the name of input and
	where it is saved.
	Be sure to check data compilation for correct data before using the script.
TODO:
	Improve Documentation
	Implement more dynamic method of grabbing data so changing order of data won't matter anymore(Check row by name)
"""

#wb2 contains the excel sheet with all data.
wb2 = load_workbook('dataCompilation.xlsx')
def grab_values(data_sheet, cell_actual, cell_numAct, cell_numExp):
	"""
	Grabs values from data_sheet to be used to show the values being used for extrapolation

	Args:
		data_sheet(worksheet): The worksheet that contains the data
		cell_actual(str): The cell that contains actual recorded value from database.
		cell_numAct(str): The cell that contains actual # of data points counted from database.
		cell_numExp(str): The cell that contains actual # of data points expected for that month
		
	Returns:
		Three variables containing the values of the requested cells.
	"""
	actual = data_sheet[cell_actual].value
	num_actual = data_sheet[cell_numAct].value
	num_expected = data_sheet[cell_numExp].value
	return actual, num_actual, num_expected

def calc_rates():
	"""
	Calculates the rates from the data_sheet, and returns the values in equation form for double checking.
	
	Args:
		None
		
	Returns:
		Three variables containing the equations for kwh, water, and gas rates
	"""
	#KWH Rate
	data_sheet = wb2.worksheets[4]
	elCapKWH = data_sheet['B1'].value
	elCapPrice = data_sheet['B2'].value
	wilsonKWH = data_sheet['B3'].value
	wilsonPrice = data_sheet['B4'].value
	sunpowerKWH = data_sheet['B5'].value
	sunpowerPrice = data_sheet['B6'].value
	UCOP = data_sheet['B7'].value
	kwhRate = '=('+repr(elCapPrice)+'+'+repr(wilsonPrice)+'+'+repr(sunpowerPrice)+'+'+repr(UCOP)+')/('+repr(elCapKWH)+'+'+repr(wilsonKWH)+'+'+repr(sunpowerKWH)+')'
	#Water Rate
	h2oPrice = data_sheet['B8'].value
	h2oHCF = data_sheet['B9'].value
	waterRate = '='+repr(h2oPrice)+'/(' +repr(h2oHCF)+'*(748/1000))'
	#Gas Rate
	gasPrice = data_sheet['B10'].value
	gasTherms = data_sheet['B11'].value
	gasRate = '='+repr(gasPrice)+'/' +repr(gasTherms)
	return kwhRate, waterRate, gasRate

def get_columns():
	"""
	Finds which column to input for all recharges. Uses a while loop so that it can handle invalid inputs.
	
	Args:
		None
		
	Returns:
		int containing column number to input all data for recharges
	"""
	#If statements to handle months that are out of order. Columns C(july)-N(june)
	while True:
		month=int(input("Input Month #(1-12): "))
		if month > 12 or month < 1:
			error = 'Invalid Month, please input a month from 1-12'
			str(error)
		elif month >= 7:
			inputColumn = 3+(month-7)
			break
		else:
			inputColumn = 3+month+5
			break
	return inputColumn
def fill_utility(inputColumn, kwhRate, waterRate, gasRate):
	"""
	Fills in data for Gall R&W, Facilities, and the Library.
	
	Args:
		inputColumn(int): Determines which column(month) is being filled in.
		kwhRate(str): The kwhrate to be filled in the recharge
		waterRate(str): The kwhrate to be filled in the recharge
		gasRate(str): The kwhrate to be filled in the recharge
		
	Returns:
		None
	"""
	wb= load_workbook('test1.xlsx')
	#Editing Gall R&W Summary
	edit_sheet = wb["Gall R&W Utility Summary"]
	#Electricity Input
	data_sheet = wb2.worksheets[0]
	actual, num_actual, num_expected = grab_values(data_sheet,'C2','E2','F2')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=7, column=inputColumn).value=input
	edit_sheet.cell(row=8, column=inputColumn).value=kwhRate

	#Gas Input
	data_sheet = wb2.worksheets[1]
	reading = data_sheet['C11'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=12, column=inputColumn).value=reading
	edit_sheet.cell(row=14, column=inputColumn).value=gasRate

	#CHW input
	data_sheet = wb2.worksheets[2]
	actual, num_actual, num_expected = grab_values(data_sheet,'C2','E2','F2')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=30, column=inputColumn).value=input

	#Water Input
	data_sheet = wb2.worksheets[3]
	edit_sheet.cell(row=20, column=inputColumn).value=waterRate
	edit_sheet.cell(row=26, column=inputColumn).value=waterRate
	reading = data_sheet['C39'].value+data_sheet['C40'].value-data_sheet['D39'].value-data_sheet['D40'].value
	edit_sheet.cell(row=19, column=inputColumn).value=reading
	reading1 = data_sheet['C61'].value-data_sheet['D61'].value
	reading2 = data_sheet['C62'].value-data_sheet['D62'].value
	edit_sheet.cell(row=24, column=inputColumn).value=reading1
	edit_sheet.cell(row=25, column=inputColumn).value=reading2
	
	#Editing Facilities Summary
	edit_sheet = wb["Facilities Utility Summary"]
	#Electricity Input
	data_sheet = wb2.worksheets[0]
	actual, num_actual, num_expected = grab_values(data_sheet,'C3','E3','F3')
	input = 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=6, column=inputColumn).value=input
	edit_sheet.cell(row=7, column=inputColumn).value=kwhRate

	#Water Input
	data_sheet = wb2.worksheets[3]
	edit_sheet.cell(row=17, column=inputColumn).value=waterRate
	reading = data_sheet['C26'].value+data_sheet['C27'].value-data_sheet['D26'].value-data_sheet['D27'].value
	edit_sheet.cell(row=16, column=inputColumn).value=reading

	#CHW input
	data_sheet = wb2.worksheets[2]
	actual, num_actual, num_expected = grab_values(data_sheet,'C3','E3','F3')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=21, column=inputColumn).value=input
	
	#Editing Library Utility Summary
	edit_sheet = wb["Library Utility Summary"]
	#Electricity Input
	data_sheet = wb2.worksheets[0]
	actual, num_actual, num_expected = grab_values(data_sheet,'C4','E4','F4')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=6, column=inputColumn).value=input
	edit_sheet.cell(row=7, column=inputColumn).value=kwhRate

	#Gas Input
	data_sheet = wb2.worksheets[1]
	actual, num_actual, num_expected = grab_values(data_sheet,'C15','E15','F15')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=11, column=inputColumn).value=input
	edit_sheet.cell(row=12, column=inputColumn).value=gasRate

	#CHW input
	data_sheet = wb2.worksheets[2]
	actual, num_actual, num_expected = grab_values(data_sheet,'C4','E4','F4')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=21, column=inputColumn).value=input

	#Water Input
	data_sheet = wb2.worksheets[3]
	edit_sheet.cell(row=17, column=inputColumn).value=waterRate
	reading = data_sheet['C34'].value+data_sheet['C35'].value-data_sheet['D34'].value-data_sheet['D35'].value
	edit_sheet.cell(row=16, column=inputColumn).value=reading

	#Save Changes to Utilities excel sheet
	wb.save('test1.xlsx')

def fill_dining(inputColumn, kwhRate, waterRate, gasRate):
	"""
	Fills in data for Dining, Laundry
	
	Args:
		inputColumn(int): Determines which column(month) is being filled in.
		kwhRate(str): The kwhrate to be filled in the recharge
		waterRate(str): The kwhrate to be filled in the recharge
		gasRate(str): The kwhrate to be filled in the recharge
		
	Returns:
		None
	"""
	#Dining Utility Summary Fill out
	wb = load_workbook('test2.xlsx')
	#Electricity Input
	edit_sheet = wb["Electricity"]
	data_sheet = wb2.worksheets[0]

	#Dining Commons KWH
	edit_sheet.cell(row=5, column=inputColumn).value=kwhRate
	actual, num_actual, num_expected = grab_values(data_sheet,'C6','E6','F6')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=8, column=inputColumn).value=input

	#Dining Expansion KWH
	actual, num_actual, num_expected = grab_values(data_sheet,'C5','E5','F5')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=11, column=inputColumn).value=input

	#Gas Input
	edit_sheet = wb["Gas"]
	data_sheet = wb2.worksheets[1]

	#Dining Boiler
	reading = data_sheet['C6'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=7, column=inputColumn).value=reading

	#Laundry
	reading = data_sheet['C7'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=11, column=inputColumn).value=reading

	#CHW Input
	edit_sheet = wb["Chilled Water"]
	data_sheet = wb2.worksheets[2]
	actual, num_actual, num_expected = grab_values(data_sheet,'C5','E5','F5')
	input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	edit_sheet.cell(row=7, column=inputColumn).value=input

	#Water Input
	edit_sheet = wb["Water"]
	data_sheet = wb2.worksheets[3] 
	edit_sheet.cell(row=5, column=inputColumn).value=waterRate
	reading = data_sheet['C34'].value+data_sheet['C35'].value-data_sheet['D34'].value-data_sheet['D35'].value
	edit_sheet.cell(row=8, column=inputColumn).value=reading

	#Saving Changes to Dining
	wb.save('test2.xlsx')
	
def fill_housing(inputColumn, kwhRate, waterRate, gasRate):
	"""
	Fills in recharges for Sierra Terraces, Valley Terraces, and Summits
	
	Args:
		inputColumn(int): Determines which column(month) is being filled in.
		kwhRate(str): The kwhrate to be filled in the recharge
		waterRate(str): The kwhrate to be filled in the recharge
		gasRate(str): The kwhrate to be filled in the recharge
		
	Returns:
		None
	"""
	#Housing Utility Summary Fill out
	wb= load_workbook('test3.xlsx')
	#Electricity Input
	edit_sheet = wb["Electricity"]
	data_sheet = wb2.worksheets[0]
	edit_sheet.cell(row=5, column=inputColumn).value=kwhRate
	for i in range(1,6):
		actual = data_sheet.cell(row=i+6, column=3).value
		num_actual = data_sheet.cell(row=i+6, column=5).value
		num_expected = data_sheet.cell(row=i+6, column=6).value
		input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
		edit_sheet.cell(row=3+i*3, column=inputColumn).value=input
	actual = data_sheet.cell(row=12, column=3).value
	edit_sheet.cell(row=21, column=inputColumn).value=actual
	for i in range(7,15):
		actual = data_sheet.cell(row=i+6, column=3).value
		num_actual = data_sheet.cell(row=i+6, column=5).value
		num_expected = data_sheet.cell(row=i+6, column=6).value
		input = '='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
		edit_sheet.cell(row=4+i*3, column=inputColumn).value=input

	#Gas Input
	edit_sheet = wb["Gas"]
	data_sheet = wb2.worksheets[1]
	edit_sheet.cell(row=5, column=inputColumn).value=gasRate

	#Dining
	reading = data_sheet['C6'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=7, column=inputColumn).value=reading

	#Laundry
	reading = data_sheet['C7'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=11, column=inputColumn).value=reading

	#Sierra Terraces
	reading = data_sheet['C10'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=15, column=inputColumn).value=reading

	#Tenaya
	reading = data_sheet['C4'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=20, column=inputColumn).value=reading

	#Cathedral
	reading = data_sheet['C3'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=25, column=inputColumn).value=reading

	#Half Dome
	reading = data_sheet['C5'].value
	reading = re.sub('cuft.', '', reading)
	edit_sheet.cell(row=30, column=inputColumn).value=reading

	#Water Input
	data_sheet = wb2.worksheets[3]
	edit_sheet = wb["Water"]
	edit_sheet.cell(row=5, column=inputColumn).value=waterRate

	#Terrace Center Commons
	reading = data_sheet['C68'].value+data_sheet['C69'].value-data_sheet['D68'].value-data_sheet['D69'].value
	edit_sheet.cell(row=8, column=inputColumn).value=reading

	#Kern
	reading = data_sheet['C32'].value+data_sheet['C33'].value-data_sheet['D32'].value-data_sheet['D33'].value
	edit_sheet.cell(row=12, column=inputColumn).value=reading

	#Tulare
	reading = data_sheet['C70'].value+data_sheet['C71'].value-data_sheet['D70'].value-data_sheet['D71'].value
	edit_sheet.cell(row=16, column=inputColumn).value=reading

	#Madera+Fresno+Stan+Kings
	reading = data_sheet['C57'].value+data_sheet['C58'].value-data_sheet['D57'].value-data_sheet['D58'].value
	reading = reading/4
	for i in range(1,5):
		edit_sheet.cell(row=15+i*4, column=inputColumn).value=reading

	#San Joaquin
	reading = data_sheet['C46'].value+data_sheet['C47'].value-data_sheet['D46'].value-data_sheet['D47'].value
	edit_sheet.cell(row=35, column=inputColumn).value=reading

	#Merced+Calaveras
	reading = data_sheet['C4'].value+data_sheet['C5'].value-data_sheet['D4'].value-data_sheet['D5'].value
	reading = reading/2
	edit_sheet.cell(row=38, column=inputColumn).value=reading
	edit_sheet.cell(row=41, column=inputColumn).value=reading

	#Sierra Terraces
	reading = data_sheet['C72'].value+data_sheet['C73'].value-data_sheet['D72'].value-data_sheet['D73'].value
	edit_sheet.cell(row=44, column=inputColumn).value=reading

	#Tenaya
	reading = data_sheet['C66'].value+data_sheet['C67'].value-data_sheet['D66'].value-data_sheet['D67'].value
	edit_sheet.cell(row=47, column=inputColumn).value=reading

	#Cathedral
	reading = data_sheet['C10'].value+data_sheet['C11'].value-data_sheet['D10'].value-data_sheet['D11'].value
	edit_sheet.cell(row=50, column=inputColumn).value=reading

	#Half Dome
	reading = data_sheet['C28'].value+data_sheet['C29'].value-data_sheet['D28'].value-data_sheet['D29'].value
	edit_sheet.cell(row=53, column=inputColumn).value=reading

	#CHW input
	edit_sheet = wb["Chilled Water"]
	data_sheet = wb2.worksheets[2]
	for i in range(1,15):
		actual=data_sheet.cell(row=i+5, column=3).value
		num_actual=data_sheet.cell(row=i+5, column=5).value
		num_expected=data_sheet.cell(row=i+5, column=6).value
		input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
		edit_sheet.cell(row=3+i*3, column=inputColumn).value=input
	#Saving Changes to Housing excel sheet
	wb.save('test3.xlsx')
	
def main():
	wb2 = load_workbook('dataCompilation.xlsx')
	inputColumn=get_columns()
	kwhRate, waterRate, gasRate=calc_rates()
	fill_utility(inputColumn,kwhRate, waterRate, gasRate)
	fill_dining(inputColumn, kwhRate, waterRate, gasRate)
	fill_housing(inputColumn, kwhRate, waterRate, gasRate)
	

if __name__== "__main__":
	main()
	
