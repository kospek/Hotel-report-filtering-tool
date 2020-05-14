import openpyxl
import os
import pandas
import time
from datetime import timedelta, date

#calculating age
def calculateAge(born):
    today = date.today()
    try:
        birthday = born.replace(year = today.year)

    # raised when birth date is February 29
    # and the current year is not a leap year
    except ValueError:
        birthday = born.replace(year = today.year,
                  month = born.month + 1, day = 1)

    if birthday > today:
        return today.year - born.year - 1
    else:
        return today.year - born.year


print('Enter path to file...')
path = input()
fldr = os.path.split(path)[0]
try:
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
except Exception as e:
    print('Incorrect path, restart program and enter valid path.')
    time.sleep(4)
    quit()


#needed columns: B, C, L, M, N, P, BE

columns = ['B', 'C', 'L', 'M', 'N', 'P', 'BE']

Name = [] #B
Surname = [] #C
DateArr = [] #L
DateLeave = [] #M
NightCount = [] #N
DateOfBirth = [] #P
CityTaxPaid = [] #BE
#record all data into lists
for col in columns:
    for cell in sheet[col]:
        if col == 'B':
            Name.append(cell.value)
        elif col == 'C':
            Surname.append(cell.value)
        elif col == 'L':
            DateArr.append(cell.value)
        elif col == 'M':
            DateLeave.append(cell.value)
        elif col == 'N':
            NightCount.append(cell.value)
        elif col == 'P':
            DateOfBirth.append(cell.value)
        elif col == 'BE':
            CityTaxPaid.append(cell.value)
#create dataframe with all data
NewTable = [Name[1:], Surname[1:], DateArr[1:], DateLeave[1:], NightCount[1:], DateOfBirth[1:], CityTaxPaid[1:]]
df_Transposed = pandas.DataFrame(NewTable).T
#name columns
df_Transposed.columns=[Name[0], Surname[0], DateArr[0], DateLeave[0], NightCount[0], DateOfBirth[0], CityTaxPaid[0]]
df_Transposed = df_Transposed.sort_values(by=[DateArr[0],DateLeave[0],Surname[0]])

#drop from table people younger 15 and older 62.
for index, col in df_Transposed.iterrows():
    if calculateAge(col[DateOfBirth[0]].dt.round('D'))<=15 or calculateAge(col[DateOfBirth[0]].dt.round('D')) >=62:
        df_Transposed = df_Transposed.drop([index])

#reset index for clearer view
df_Transposed = df_Transposed.reset_index(drop=True)

#change date format to dd.mm.yyyy. NOTE: this will change object type from "date" to "string"
df_Transposed[DateArr[0]] = df_Transposed[DateArr[0]].dt.strftime('%d.%m.%Y')
df_Transposed[DateLeave[0]] = df_Transposed[DateLeave[0]].dt.strftime('%d.%m.%Y')
df_Transposed[DateOfBirth[0]] = df_Transposed[DateOfBirth[0]].dt.strftime('%d.%m.%Y')

#export
exp_pth = os.path.join(fldr, 'sorted.xlsx')
df_Transposed.to_excel(exp_pth)
print('Sorted file is saved as:', exp_pth)
print('press enter to quit')
input()
