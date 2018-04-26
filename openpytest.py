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
#Gall R&W Utility Summary, Facilities ... ..., Library ... ...
#Currently Cannot think of way to easily iterate, maybe future find a way.
wb= load_workbook('test.xlsx')

#Editing Gall R&W Summary
editSheet = wb["Gall R&W Utility Summary"]

#Electricity Input
dataSheet = wb2.worksheets[0]
actual=dataSheet['C2'].value
num_actual=dataSheet['E2'].value
num_expected=dataSheet['F2'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=7, column=inputColumn).value=input
editSheet.cell(row=8, column=inputColumn).value=kwhRate
#Gas Input
dataSheet = wb2.worksheets[1]
reading=dataSheet['C11'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=12, column=inputColumn).value=reading
editSheet.cell(row=13, column=inputColumn).value=gasRate
#CHW input
dataSheet = wb2.worksheets[2]
actual=dataSheet['C2'].value
num_actual=dataSheet['E2'].value
num_expected=dataSheet['F2'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=30, column=inputColumn).value=input

#Editing Facilities Summary
editSheet = wb["Facilities Utility Summary"]
#Electricity Input
dataSheet = wb2.worksheets[0]
actual=dataSheet['C3'].value
num_actual=dataSheet['E3'].value
num_expected=dataSheet['F3'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=6, column=inputColumn).value=input
editSheet.cell(row=7, column=inputColumn).value=kwhRate
#CHW input
dataSheet = wb2.worksheets[2]
actual=dataSheet['C3'].value
num_actual=dataSheet['E3'].value
num_expected=dataSheet['F3'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=21, column=inputColumn).value=input

#Editing Library Utility Summary
editSheet = wb["Library Utility Summary"]
#Electricity Input
dataSheet = wb2.worksheets[0]
actual=dataSheet['C4'].value
num_actual=dataSheet['E4'].value
num_expected=dataSheet['F4'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=6, column=inputColumn).value=input
editSheet.cell(row=7, column=inputColumn).value=kwhRate
#Gas Input
dataSheet = wb2.worksheets[1]
actual=dataSheet['C15'].value
num_actual=dataSheet['E15'].value
num_expected=dataSheet['F15'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=11, column=inputColumn).value=input
editSheet.cell(row=12, column=inputColumn).value=input
#CHW input
dataSheet = wb2.worksheets[2]
actual=dataSheet['C4'].value
num_actual=dataSheet['E4'].value
num_expected=dataSheet['F4'].value
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=21, column=inputColumn).value=input

#TODO: The other reacharge spreadsheets baby
#Example code of inputing into specific cell  d = ws.cell(row=4, column=2, value=10)
#Ask Gabriel to format Gas reads to be like Utility Recharges (Months in same columns)

#Save Changes to chosen excel sheet
wb.save('test.xlsx')
