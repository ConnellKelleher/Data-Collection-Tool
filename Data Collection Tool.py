import xlrd
import math

totalCust = 0
totalSaved = 0
totalOwner = 0
totalOwnerSaved = 0
wb = xlrd.open_workbook('report2.xls')
sheet = wb.sheet_by_index(0)
total = 0

#Gets total number of rows in excel sheet
for i in range(sheet.nrows):
    total = total + 1

#checks if each transaction is from an Owner or Non-Owner
#Adds non-owners and owners and totals the amount saved.
for i in range(sheet.nrows):
    if sheet.cell_value(i,13) == 'Cashier':
        for i in range(i, i + 25):
            if sheet.cell_value(i,4) == 'Mfr. Coupon 08600063544-021122':
                totalOwner = totalOwner + 1
                totalOwnerSaved = totalOwnerSaved + sheet.cell_value(i,21)
            if i == total - 1:
                break
    elif sheet.cell_value(i,13) == 'None':
        for i in range(i, i + 25):
            if sheet.cell_value(i,4) == 'Mfr. Coupon 08600063544-021122':
                totalCust = totalCust + 1
                totalSaved = totalSaved + sheet.cell_value(i,21)
            if i == total - 1:
               break

#Rounded Total Saved To Two Decimal Places
roundedSaved = round(totalSaved, 2)
roundedOwnerSaved = round(totalOwnerSaved, 2)

#Prints Total Owners and Total Saved
print("The total number of coupons used by owners was " + str(totalOwner))
print("The total number of coupons used by non-owners was " + str(totalCust))
print("The total amount that Owners saved was using their discount was $" + str(roundedOwnerSaved))
print("The total amount that Non-Owners saved was using their discount was $" + str(roundedSaved))

#Mfr. Coupon 08600063544-021122
