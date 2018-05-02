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
# Columns C(july)-N(june) Use input month # to determine which column to input data.
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

#Grabs the actual value/(num actual/num expected) for the equation.
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
wb= load_workbook('02-2018_Utility_Summary.xlsx')

#Editing Gall R&W Summary
editSheet = wb["Gall R&W Utility Summary"]
#Electricity Input
dataSheet = wb2.worksheets[0]
actual, num_actual, num_expected = grab_values('C2','E2','F2')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=7, column=inputColumn).value=input
editSheet.cell(row=8, column=inputColumn).value=kwhRate

#Gas Input
dataSheet = wb2.worksheets[1]
reading=dataSheet['C11'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=12, column=inputColumn).value=reading
editSheet.cell(row=14, column=inputColumn).value=gasRate

#CHW input
dataSheet = wb2.worksheets[2]
actual, num_actual, num_expected = grab_values('C2','E2','F2')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=30, column=inputColumn).value=input

#Water Input
dataSheet = wb2.worksheets[3]
editSheet.cell(row=20, column=inputColumn).value=waterRate
editSheet.cell(row=26, column=inputColumn).value=waterRate
reading=dataSheet['C39'].value+dataSheet['C40'].value-dataSheet['D39'].value-dataSheet['D40'].value
editSheet.cell(row=19, column=inputColumn).value=reading
reading1=dataSheet['C61'].value-dataSheet['D61'].value
reading2=dataSheet['C62'].value-dataSheet['D62'].value
editSheet.cell(row=24, column=inputColumn).value=reading1
editSheet.cell(row=25, column=inputColumn).value=reading2

#Editing Facilities Summary
editSheet = wb["Facilities Utility Summary"]
#Electricity Input
dataSheet = wb2.worksheets[0]
actual, num_actual, num_expected = grab_values('C3','E3','F3')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=6, column=inputColumn).value=input
editSheet.cell(row=7, column=inputColumn).value=kwhRate

#Water Input
dataSheet = wb2.worksheets[3]
editSheet.cell(row=17, column=inputColumn).value=waterRate
reading=dataSheet['C26'].value+dataSheet['C27'].value-dataSheet['D26'].value-dataSheet['D27'].value
editSheet.cell(row=16, column=inputColumn).value=reading

#CHW input
dataSheet = wb2.worksheets[2]
actual, num_actual, num_expected = grab_values('C3','E3','F3')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=21, column=inputColumn).value=input

#Editing Library Utility Summary
editSheet = wb["Library Utility Summary"]
#Electricity Input
dataSheet = wb2.worksheets[0]
actual, num_actual, num_expected = grab_values('C4','E4','F4')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=6, column=inputColumn).value=input
editSheet.cell(row=7, column=inputColumn).value=kwhRate

#Gas Input
dataSheet = wb2.worksheets[1]
actual, num_actual, num_expected = grab_values('C15','E15','F15')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=11, column=inputColumn).value=input
editSheet.cell(row=12, column=inputColumn).value=gasRate

#CHW input
dataSheet = wb2.worksheets[2]
actual, num_actual, num_expected = grab_values('C4','E4','F4')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=21, column=inputColumn).value=input

#Water Input
dataSheet = wb2.worksheets[3]
editSheet.cell(row=17, column=inputColumn).value=waterRate
reading=dataSheet['C34'].value+dataSheet['C35'].value-dataSheet['D34'].value-dataSheet['D35'].value
editSheet.cell(row=16, column=inputColumn).value=reading

#Save Changes to Utilities excel sheet
wb.save('02-2018_Utility_Summary.xlsx')

#Housing Utility Summary Fill out
wb= load_workbook('02-2018_Housing_Utility_Summary.xlsx')
#Electricity Input
editSheet = wb["Electricity"]
dataSheet = wb2.worksheets[0]
editSheet.cell(row=5, column=inputColumn).value=kwhRate
for i in range(1,6):
	actual=dataSheet.cell(row=i+6, column=3).value
	num_actual=dataSheet.cell(row=i+6, column=5).value
	num_expected=dataSheet.cell(row=i+6, column=6).value
	input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	editSheet.cell(row=3+i*3, column=inputColumn).value=input
actual=dataSheet.cell(row=12, column=3).value
editSheet.cell(row=21, column=inputColumn).value=actual
for i in range(7,15):
	actual=dataSheet.cell(row=i+6, column=3).value
	num_actual=dataSheet.cell(row=i+6, column=5).value
	num_expected=dataSheet.cell(row=i+6, column=6).value
	input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	editSheet.cell(row=4+i*3, column=inputColumn).value=input

#Gas Input
editSheet = wb["Gas"]
dataSheet = wb2.worksheets[1]
editSheet.cell(row=5, column=inputColumn).value=gasRate

#Dining
reading= dataSheet['C6'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=7, column=inputColumn).value=reading

#Laundry
reading= dataSheet['C7'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=11, column=inputColumn).value=reading

#Sierra Terraces
reading= dataSheet['C10'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=15, column=inputColumn).value=reading

#Tenaya
reading= dataSheet['C4'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=20, column=inputColumn).value=reading

#Cathedral
reading= dataSheet['C3'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=25, column=inputColumn).value=reading

#Half Dome
reading= dataSheet['C5'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=30, column=inputColumn).value=reading

#Water Input
dataSheet = wb2.worksheets[3]
editSheet = wb["Water"]
editSheet.cell(row=5, column=inputColumn).value=waterRate

#Terrace Center Commons
reading=dataSheet['C68'].value+dataSheet['C69'].value-dataSheet['D68'].value-dataSheet['D69'].value
editSheet.cell(row=8, column=inputColumn).value=reading

#Kern
reading=dataSheet['C32'].value+dataSheet['C33'].value-dataSheet['D32'].value-dataSheet['D33'].value
editSheet.cell(row=12, column=inputColumn).value=reading

#Tulare
reading=dataSheet['C70'].value+dataSheet['C71'].value-dataSheet['D70'].value-dataSheet['D71'].value
editSheet.cell(row=16, column=inputColumn).value=reading

#Madera+Fresno+Stan+Kings
reading=dataSheet['C57'].value+dataSheet['C58'].value-dataSheet['D57'].value-dataSheet['D58'].value
reading= reading/4
for i in range(1,5):
	editSheet.cell(row=15+i*4, column=inputColumn).value=reading

#San Joaquin
reading=dataSheet['C46'].value+dataSheet['C47'].value-dataSheet['D46'].value-dataSheet['D47'].value
editSheet.cell(row=35, column=inputColumn).value=reading

#Merced+Calaveras
reading=dataSheet['C4'].value+dataSheet['C5'].value-dataSheet['D4'].value-dataSheet['D5'].value
reading=reading/2
editSheet.cell(row=38, column=inputColumn).value=reading
editSheet.cell(row=41, column=inputColumn).value=reading

#Sierra Terraces
reading=dataSheet['C72'].value+dataSheet['C73'].value-dataSheet['D72'].value-dataSheet['D73'].value
editSheet.cell(row=44, column=inputColumn).value=reading

#Tenaya
reading=dataSheet['C66'].value+dataSheet['C67'].value-dataSheet['D66'].value-dataSheet['D67'].value
editSheet.cell(row=47, column=inputColumn).value=reading

#Cathedral
reading=dataSheet['C10'].value+dataSheet['C11'].value-dataSheet['D10'].value-dataSheet['D11'].value
editSheet.cell(row=50, column=inputColumn).value=reading

#Half Dome
reading=dataSheet['C28'].value+dataSheet['C29'].value-dataSheet['D28'].value-dataSheet['D29'].value
editSheet.cell(row=53, column=inputColumn).value=reading

#CHW input
editSheet = wb["Chilled Water"]
dataSheet = wb2.worksheets[2]
for i in range(1,15):
	actual=dataSheet.cell(row=i+5, column=3).value
	num_actual=dataSheet.cell(row=i+5, column=5).value
	num_expected=dataSheet.cell(row=i+5, column=6).value
	input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
	editSheet.cell(row=3+i*3, column=inputColumn).value=input
#Saving Changes to Housing excel sheet
wb.save('02-2018_Housing_Utility_Summary.xlsx')

#Dining Utility Summary Fill out
wb= load_workbook('02-2018_Dining_Utility_Summary.xlsx')
#Electricity Input
editSheet = wb["Electricity"]
dataSheet = wb2.worksheets[0]

#Dining Commons KWH
editSheet.cell(row=5, column=inputColumn).value=kwhRate
actual, num_actual, num_expected = grab_values('C6','E6','F6')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=8, column=inputColumn).value=input

#Dining Expansion KWH
actual, num_actual, num_expected = grab_values('C5','E5','F5')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=11, column=inputColumn).value=input

#Gas Input
editSheet = wb["Gas"]
dataSheet = wb2.worksheets[1]

#Dining Boiler
reading= dataSheet['C6'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=7, column=inputColumn).value=reading

#Laundry
reading= dataSheet['C7'].value
reading = re.sub('cuft.', '', reading)
editSheet.cell(row=11, column=inputColumn).value=reading

#CHW Input
editSheet = wb["Chilled Water"]
dataSheet = wb2.worksheets[2]
actual, num_actual, num_expected = grab_values('C5','E5','F5')
input= 	'='+repr(actual)+'/('+repr(num_actual)+'/'+repr(num_expected)+')'
editSheet.cell(row=7, column=inputColumn).value=input

#Water Input
editSheet = wb["Water"]
dataSheet = wb2.worksheets[3]
editSheet.cell(row=5, column=inputColumn).value=waterRate
reading=dataSheet['C34'].value+dataSheet['C35'].value-dataSheet['D34'].value-dataSheet['D35'].value
editSheet.cell(row=8, column=inputColumn).value=reading

#Saving Changes to Dining
wb.save('02-2018_Dining_Utility_Summary.xlsx')












