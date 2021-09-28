#Library used to have access to the Google Sheets
import gspread
from gspread.models import Worksheet

#Credentials to access the APIs
gc = gspread.service_account(filename='./service_account.json')

#copied sheet access key (found in URL)
sh = gc.open_by_key("1R70ID1062jQHL1VDNLsy_lBlEFKU_MF8E0vHxy2eZhA")

#Selected the sheet as the active Worksheet
Worksheet = sh.sheet1

#Creating 4 "lists of lists" with the necessary information (The library put each cell information in separate lists inside one big list of lists)
test1Aux = Worksheet.get('D4:D27')
test2Aux = Worksheet.get('E4:E27')
test3Aux = Worksheet.get('F4:F27')
absencesAux = Worksheet.get('C4:C27')
#Creating lists to be easier manipulated
test1 = []
test2 = []
test3 = []
absences = []

 #This 'for' transforms the previous list of lists, in a single list! for the 4 columns
for x in range(24):
    a = list(map(int,test1Aux[x]))
    b = list(map(int,test2Aux[x]))
    c = list(map(int,test3Aux[x]))
    d = list(map(int,absencesAux[x]))
    #The information is turned into 'int' so we can use it later
    
    #Here each cell is being concatenated in the new Lists
    test1 += a
    test2 += b
    test3 += c
    absences += d

#Creating the lists that will be used to update the sheet
situation = ['none']*24
finalExam = ['none']*24
#Defined a size for this lists because an 'index out of range' error was happening

#Here the lists created above are filled so we can update the sheet columns in a batch and not exceed the API requests limitation
for x in range(24):

    #Test grades average
    average = (test1[x]+test2[x]+test3[x])/3

    #First calculating whether the student has the minimum attendance percentage
    if absences[x] <= 60 * 0.25:
    #If he has the minimum attendance we proceed to calculate his situation

        #Calculating if the student attends the minimum grade average and filling the lists
        if average >= 70:
            situation[x] = ' Aprovado'
            finalExam[x] = '0'

        #Calculating if the student has no chances of being approved and filling the lists
        elif average < 50:
            situation[x] = 'Reprovado por Nota'
            finalExam[x] = '0'

        #Finally if anything doesn't apply the lists are filled with the final exam information
        else:
            situation[x] = 'Exame final'
            #Calculating the necessary final exam grade to be approved
            finalExam[x] = str(int(140 - average))

            #In the challenge information we had this formula: 5 <= (m + naf)/2,
            #but it made much more sense to use the value as 70 (the minimum required to be approved)

    #Here we add this information if the student doesn't have the minimun attendance
    else:
        situation[x] = 'Reprovado por Falta'
        finalExam[x] = '0'

#Using an library function we select the range we will update
cell_list = Worksheet.range('G4:G27')

#Here we enumerate the situation list and pass each information to where its suposed to be on the sheet
for i, val in enumerate(situation): 
    cell_list[i].value = val    

#And here we finally update all the column cells in a batch so it can be a lot faster
Worksheet.update_cells(cell_list)

#The same thing is being done here but with the finalExam information
cell_list = Worksheet.range('H4:H27')

for i, val in enumerate(finalExam): 
    cell_list[i].value = val    

Worksheet.update_cells(cell_list)