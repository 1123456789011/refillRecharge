from openpyxl import load_workbook
import re
#Instructions
#This script will move all compiled data and inputed rates into recharges
#In order to run this script, you must first have a worksheet with the compiled data.
#Export from the data base the CHW and KWH data (be sure to have column names)
#Fill in data from reads for Water and Gas
#Correct any anomalies in the compiled data.
#Uses openpyxl library for excel sheet manipulation
#Uses re for easier string removal
#wb2 contains the excel sheet with all data. Be sure to retain format to ensure script works
wb2 = load_workbook('dataCompilation.xlsx')
# C(july)-N(june) Use input month # to determine which column to input data.
while True:
	month=int(input("Input Month #(1-12): "))
	if month > 12 or month < 1:
		error= 'Invalid Month, please input a month from 1-12'
		str(error)
	elif month >= 7:
		inputColumn= 3+(month-7)
		break
	else:
		inputColumn= 3+month+5
		break

def grab_values(cell_actual, cell_numAct, cell_numExp):
	actual=dataSheet[cell_actual].value
	num_actual=dataSheet[cell_numAct].value
	num_expected=dataSheet[cell_numExp].value
	return actual, num_actual, num_expected
#KWH Rate
dataSheet = wb2.worksheets[4]
elCapKWH= dataSheet['B1'].value
elCapPrice=dataSheet['B2'].value
wilsonKWH=dataSheet['B3'].value
wilsonPrice=dataSheet['B4'].value
sunpowerKWH=dataSheet['B5'].value
sunpowerPrice=dataSheet['B6'].value
UCOP=dataSheet['B7'].value
kwhRate= '=('+repr(elCapPrice)+'+'+repr(wilsonPrice)+'+'+repr(sunpowerPrice)+'+'+repr(UCOP)+')/('+repr(elCapKWH)+'+'+repr(wilsonKWH)+'+'+repr(sunpowerKWH)+')'
#Water Rate
h2oPrice=dataSheet['B8'].value
h2oHCF=dataSheet['B9'].value
waterRate= '='+repr(h2oPrice)+'/(' +repr(h2oHCF)+'*(748/1000))'
#Gas Rate
gasPrice=dataSheet['B10'].value
gasTherms=dataSheet['B11'].value
gasRate= '='+repr(gasPrice)+'/' +repr(gasTherms)

#Inputing Utility_Summary Code, three sheets to edit
#Currently Cannot think of way to easily iterate, maybe future find a way.
# wb= load_workbook('test.xlsx')

# #Editing Gall R&W Summary
# editSheet = wb["Gall R&W Utility Summary"]

# #Electricity Input
# dataSheet = wb2.worksheets[0]
# actual=dataSheet['C2'].value
# num_actual=dataSheet['E2'].value
# num_expected=dataSheet['F2'].value
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=7, column=inputColumn).value=input
# editSheet.cell(row=5, column=inputColumn).value=kwhRate
# #Gas Input
# dataSheet = wb2.worksheets[1]
# reading=dataSheet['C11'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=12, column=inputColumn).value=reading
# editSheet.cell(row=13, column=inputColumn).value=gasRate
# #CHW input
# dataSheet = wb2.worksheets[2]
# actual=dataSheet['C2'].value
# num_actual=dataSheet['E2'].value
# num_expected=dataSheet['F2'].value
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=30, column=inputColumn).value=input

# #Editing Facilities Summary
# editSheet = wb["Facilities Utility Summary"]
# #Electricity Input
# dataSheet = wb2.worksheets[0]
# actual, num_actual, num_expected = grab_values('C3','E3','F3')
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=6, column=inputColumn).value=input
# editSheet.cell(row=7, column=inputColumn).value=kwhRate
# #CHW input
# dataSheet = wb2.worksheets[2]
# actual, num_actual, num_expected = grab_values('C3','E3','F3')
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=21, column=inputColumn).value=input

# #Editing Library Utility Summary
# editSheet = wb["Library Utility Summary"]
# #Electricity Input
# dataSheet = wb2.worksheets[0]
# actual, num_actual, num_expected = grab_values('C4','E4','F4')
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=6, column=inputColumn).value=input
# editSheet.cell(row=7, column=inputColumn).value=kwhRate
# #Gas Input
# dataSheet = wb2.worksheets[1]
# actual, num_actual, num_expected = grab_values('C15','E15','F15')
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=11, column=inputColumn).value=input
# editSheet.cell(row=12, column=inputColumn).value=gasRate
# #CHW input
# dataSheet = wb2.worksheets[2]
# actual, num_actual, num_expected = grab_values('C4','E4','F4')
# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
# editSheet.cell(row=21, column=inputColumn).value=input

# #Save Changes to chosen excel sheet
# wb.save('test.xlsx')

# #Housing Utility Summary Fill out
# wb= load_workbook('test2b.xlsx')
# #Electricity Input
# editSheet = wb["Electricity"]
# dataSheet = wb2.worksheets[0]
# editSheet.cell(row=5, column=inputColumn).value=kwhRate
# for i in range(1,6):
	# actual=dataSheet.cell(row=i+6, column=3).value
	# num_actual=dataSheet.cell(row=i+6, column=5).value
	# num_expected=dataSheet.cell(row=i+6, column=6).value
	# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	# editSheet.cell(row=3+i*3, column=inputColumn).value=input
# actual=dataSheet.cell(row=12, column=3).value
# editSheet.cell(row=21, column=inputColumn).value=actual
# for i in range(7,15):
	# actual=dataSheet.cell(row=i+6, column=3).value
	# num_actual=dataSheet.cell(row=i+6, column=5).value
	# num_expected=dataSheet.cell(row=i+6, column=6).value
	# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	# editSheet.cell(row=4+i*3, column=inputColumn).value=input

# #Gas Input
# editSheet = wb["Gas"]
# dataSheet = wb2.worksheets[1]
# editSheet.cell(row=5, column=inputColumn).value=gasRate
# #Dining
# reading= dataSheet['C6'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=7, column=inputColumn).value=reading
# #Laundry
# reading= dataSheet['C7'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=11, column=inputColumn).value=reading
# #Sierra Terraces
# reading= dataSheet['C10'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=15, column=inputColumn).value=reading
# #Tenaya
# reading= dataSheet['C4'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=20, column=inputColumn).value=reading
# #Cathedral
# reading= dataSheet['C3'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=25, column=inputColumn).value=reading
# #Half Dome
# reading= dataSheet['C5'].value
# reading = re.sub('cuft.', '', reading)
# editSheet.cell(row=30, column=inputColumn).value=reading

# #CHW input
# editSheet = wb["Chilled Water"]
# dataSheet = wb2.worksheets[2]
# for i in range(1,15):
	# actual=dataSheet.cell(row=i+5, column=3).value
	# num_actual=dataSheet.cell(row=i+5, column=5).value
	# num_expected=dataSheet.cell(row=i+5, column=6).value
	# input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	# editSheet.cell(row=3+i*3, column=inputColumn).value=input
# wb.save('test2b.xlsx')

#Dining Utility Summary Fill out
wb= load_workbook('test3b.xlsx')
#Electricity Input
editSheet = wb["Electricity"]
dataSheet = wb2.worksheets[0]

wb.save('test3b.xlsx')
