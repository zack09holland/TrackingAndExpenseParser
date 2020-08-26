#############################################################################
#	autoTE.py
#	by Zachary Holland
#	---
#	-Drag and drop necessary pdf and excel files
#	-Process files and parse for necessary information
#	-Create a new excel file for the TE form for excavation
#	-Update the tracking form for the excavations
#		-Will need to manually enter some information as not all is obtainable
#		 through the pdfs provided
#############################################################################
# All imports are supported with 3.x
from subprocess import Popen, PIPE
import win32com.client as win32
import re
import openpyxl
import os
import time


#####################################################################
#
# Grab the Bid Amount value from the BIDDING AMOUNT PDF
#
# 	-User will drag pdf file into cmd and have it be
#  	processed into a readable text file
#
#####################################################################
fileLocation = input("Please drag pdf file loc w/ BID AMOUNT to cmd screen(press enter when done): \n")
fileLocation = fileLocation[1:-1]
process = Popen(['pdftotext', '-layout', fileLocation, 'BidAmountOutput.txt'],stdout=PIPE,stderr=PIPE)
stdout, stderr = process.communicate()
print (stderr)

#Open the text file to find the Total price
textfile = open('BidAmountOutput.txt', 'rt')
filetext = textfile.read()
textfile.close()
bidAmtMatch = re.findall('Total Price \$\s*(.*)', filetext)

print ("Completed processing of Bid PDF\n")


#####################################################################
#
# Grab the Total Amount value from the INVOICE PDF
#
# 	-User will drag pdf file into cmd and have it be
#  	processed into a readable text file
#
#####################################################################
fileLocation = input("Please drag pdf file for the INVOICE to cmd screen(press enter when done): \n")
fileLocation = fileLocation[1:-1]
process = Popen(['pdftotext', '-layout', fileLocation, 'TotalAmountOutput.txt'],stdout=PIPE,stderr=PIPE)
stdout, stderr = process.communicate()
print (stderr)

#Open the text file to find the Total price
textfile = open('TotalAmountOutput.txt', 'rt')
filetext = textfile.read()
textfile.close()
#totAmtMatch = re.findall('Total\s*\$(\S*)', filetext)
excavationCompleted = re.findall('completed (.*)',filetext)

print ("Completed processing of Invoice PDF\n")

#####################################################################
#
# Grab the necessary Excavation Request information
#
# 	-User will drag pdf file into cmd and have it be
#  	processed into a readable text file
#	-Notification Number, Work Order Number, 
#	 Address Number and Street, City Name, 
#	 Sewer Owner (Name) and Excavation Comments
#
#####################################################################
fileLocation = input("Please drag pdf of the EXCAVATION REQUEST to cmd screen(press enter when done): \n")
fileLocation = fileLocation[1:-1]
process = Popen(['pdftotext', fileLocation, 'ExRequestOutput.txt'],stdout=PIPE,stderr=PIPE)
stdout, stderr = process.communicate()
print (stderr)

#Open the text file to find necessary excavation request information
textfile = open('ExRequestOutput.txt', 'rt')
filetext = textfile.read()
textfile.close()

address = re.search('Address: (\S*\s\S*\s\S*\s\S*) (\S+)', filetext)
streetName = address.group(1)
cityName = address.group(2)

addNumSearch = re.search('Address: (\S*)', filetext)
addNum = addNumSearch.group(1)

streetNameSearch = re.search('Address: \S*(\s\S*\s\S*\s\S*)',filetext)
justStreetName = streetNameSearch.group(1)
streetName_forSaveName = streetName.replace(" ","_")

notNumber = re.findall('Notification Number: (\S*)', filetext)
workOrderNumber = re.findall('Work Order Number: (\S*)', filetext)
homeownerInfo = re.search('Homeowner info: (\S*\s\S*) Phone: (\S*)', filetext)
homeownerName = homeownerInfo.group(1)
homeownerPhone = homeownerInfo.group(2)


excavationNotes = re.findall('Additional Notes: (.+)', filetext)
print ("Completed processing of Excavation Request PDF\n")
#####################################################################
#
# Put information into excel file
#
#####################################################################
#
# Creates/Loads the workbook object for the daily timesheet
#
fileLocation = input("Please drag the T&E excel file to cmd screen(press enter when done): \n")
fileLocation = fileLocation[1:-1]
# Get the date
todaysDate = time.strftime("%m/%d/%Y")
reverseDate = time.strftime("%Y%m%d")
fileSaveName = str(notNumber[0])+'_'+str(workOrderNumber[0])+'_'+streetName_forSaveName+'_'+cityName+'_'+reverseDate
print (fileSaveName)

# Open the excel application and create a workbook object to work with
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(fileLocation)

estimateWS = wb.Worksheets('T&E Estimate')
actualWS = wb.Worksheets('T & E Actual')
#Date
estimateWS.Cells(3,3).Value = todaysDate
actualWS.Cells(3,3).Value = todaysDate
#Notification Number
estimateWS.Cells(7,3).Value = int(notNumber[0])
actualWS.Cells(8,3).Value = int(notNumber[0])
#WorkOrder Number
estimateWS.Cells(8,3).Value = int(workOrderNumber[0])
actualWS.Cells(9,3).Value = int(workOrderNumber[0])
#Street Name
estimateWS.Cells(9,3).Value = str(streetName)
actualWS.Cells(10,3).Value = str(streetName)
#City Name
estimateWS.Cells(10,3).Value = str(cityName)
actualWS.Cells(11,3).Value = str(cityName)
#Homeowner Name
estimateWS.Cells(11,3).Value = str(homeownerName)
actualWS.Cells(12,3).Value = str(homeownerName)
#ExcavationNotes
estimateWS.Cells(14,2).Value = str(excavationNotes[0])
actualWS.Cells(14,2).Value = str(excavationNotes[0])
#Bid and Total Amounts
estimateWS.Cells(24,13).Value = str(bidAmtMatch[0])
actualWS.Cells(22,13).Value = str(totAmtMatch[0])

#####################################################################
#
# Saving/Closing T&E Process Form
#
#####################################################################
# Create filepath for where workbook will be saved aka the Desktop
wbSavePath = os.path.join(r'C:\Users\Field Crew\Desktop\\'+fileSaveName+'.xlsx')
# Save the workbook with designated filename to the desktop as .xlsx(=51)
wb.SaveAs(Filename=wbSavePath, FileFormat="51")
# Close file as it is no longer needed atm, if review is wanted remove these
# statements
wb.Close(0)
excel.Quit()
print ("Completed and updated TE form\n")
#####################################################################
# 
# Update Tracking Form
#
#####################################################################
#	Status Date = (r,A)
#	AddNum = (r,E)
#	Street = (r,F)
#	City = (r,G)
#	ResidentName = (r,H)
#	ResidentPhone = (r,I)
#	BidAmount = (r,M)
#	CompletedDate = (r,O)
#	Final_Inv = (r,P)
#	TE_Total_w_Markup = (r,R)
#		=> teTotal * 1.14
#	Notes = (r,S)
#
print ("\nTracking form will now be updated...")
excel = win32.gencache.EnsureDispatch('Excel.Application')

fileLocation = input("Please drag the most up to date Tracking Form to cmd(press enter when done): \n")
fileLocation = fileLocation[1:-1]
wbTF = excel.Workbooks.Open(fileLocation)

currentMonth_Year = time.strftime("%B %Y")
currentTF_WS = wbTF.Worksheets(currentMonth_Year)


# Get last Row of the Sheet in use as well as the Column(if needed)
#lastCol = currentTF_WS.UsedRange.Columns.Count
lastRow = currentTF_WS.UsedRange.Rows.Count
currentTF_WS.Cells(lastrow+1,1).Value = todaysDate
currentTF_WS.Cells(lastRow+1,5).Value = str(addNum)
currentTF_WS.Cells(lastRow+1,6).Value = str(justStreetName)
currentTF_WS.Cells(lastRow+1,7).Value = str(cityName)
currentTF_WS.Cells(lastRow+1,8).Value = str(homeownerName)
currentTF_WS.Cells(lastRow+1,9).Value = str(homeownerPhone)
#currentTF_WS.Cells(lastRow+1,13).Value = str(bidAmtMatch[0])
currentTF_WS.Cells(lastRow+1,15).Value = str(excavationCompleted[0])
#currentTF_WS.Cells(lastRow+1,16).Value = str(totAmtMatch[0])
# In order to convert the total amount string to a float we need to 
# replace the ',' with nothing
#newFrmtTotal = str(totAmtMatch[0]).replace(',','')
#currentTF_WS.Cells(lastRow+1,18).Value = round(float(newFrmtTotal)*1.14,2)
currentTF_WS.Cells(lastRow+1,19).Value = str(excavationNotes[0])


#####################################################################
#
# Saving/Closing Tracking Form
#
#####################################################################
# Create save path for the new tracking sheet. If need be we can save to original location
# instead of to the Desktop
#wbTFSavePath = os.path.join('C:\Users\Zach-HUSA\Desktop\\'+'Excavations_Monthly_Tracking_Working.xlsx')
wbSavePath = os.path.join(r'C:\Users\Field Crew\Desktop\\'+fileSaveName+'.xlsx')

# Save the workbook with designated filename to the desktop as .xlsx(=51)
#wbTF.SaveAs(Filename=wbTFSavePath, FileFormat="51")
wbTF.SaveAs(Filename=fileLocation, FileFormat="51")

wbTF.Close(0)
excel.Quit()

print ("Finished!")
