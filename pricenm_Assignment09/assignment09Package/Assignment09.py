# assignment09.py

# Name: Nicole Price
# email: pricenm@mail.uc.edu
# Assignment Title: Assignment 09
# Course: IS 4010
# Semester/Year: Spring 2023
# Brief Description: This is an in-class assignment
# Citations: 
# Anything else that's relevant:

from openpyxl import load_workbook

def assignment09():

    wb = load_workbook(filename = 'empl.xlsx')
    sheet = wb['Sheet1']
    # print(sheet['A1'].value), print(sheet['A2'].value)
    # print(sheet['B1'].value), print(sheet['B2'].value)
    # print(sheet['C1'].value), print(sheet['C2'].value)
    
    # entire list of all cells in column C that have last names C2 to C1001
    '''
    # what i was able to come up with
    column = sheet['C'] # column
    names = [column[x].value for x in range(len(column))]
    print(names[1]) # print first cell
    print(names[-1]) # print last cell
    '''
    # what professor used to accomplish the same task
    names = [cell.value for cell in sheet['C'][1:]]
    print(names[0])
    print(names[-1])
    
    
    
    wb.close()   