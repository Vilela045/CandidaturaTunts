import gspread      #library used to have access to the Google Sheets
from gspread.models import Worksheet

gc = gspread.service_account(filename='./service_account.json')     #credentials to access the APIs

sh = gc.open_by_key("1R70ID1062jQHL1VDNLsy_lBlEFKU_MF8E0vHxy2eZhA")     #copied sheet access key (found in URL)

Worksheet = sh.sheet1

p1Aux = Worksheet.get('D4:D27')
p2Aux = Worksheet.get('E4:E27')
p3Aux = Worksheet.get('F4:F27')     #This part creates an 3 lists of lists 
p1 = []
p2 = []
p3 = []

for x in range(24):                 #This 'for' transforms the previous list of lists, in a single list! for the three columns
    a = list(map(int,p1Aux[x]))
    b = list(map(int,p2Aux[x]))     #the information is turned into 'int' so we can use it later
    c = list(map(int,p3Aux[x]))
    
    p1 = p1+a
    p2 = p2+b
    p3 = p3+c

sit = ['none']*24
final = ['none']*24         #defined a size for this lists so the 'index out of range' error wouldn't happen

for x in range(24):         #this 'for' fills 2 lists using the previous generated ones so we can update the sheet columns in a batch

    m = (p1[x]+p2[x]+p3[x])/3       #grade average

    if m >= 70:
        sit[x] = ' Aprovado'
        final[x] = '0'

    elif m < 50:
        sit[x] = 'Reprovado por Nota'
        final[x] = '0'

    else:
        sit[x] = 'Exame final'
        final[x] = str(int(140 - m))    #in the information we had this formula: 5 <= (m + naf)/2
                                        #but it made much more sense to use the value as 70 (the minimum required to be approved)
cell_list = Worksheet.range('G4:G27')
cell_values = sit

for i, val in enumerate(cell_values): 
    cell_list[i].value = val    

Worksheet.update_cells(cell_list)               #enumerated auxiliary lists and used a loop to update the range 

cell_list = Worksheet.range('H4:H27')
cell_values = final

for i, val in enumerate(cell_values): 
    cell_list[i].value = val    

Worksheet.update_cells(cell_list)