# In the command prompt, write __pip install thing____

import xlrd



def extract():
    file_location = "C:/Users/Tariq/Downloads/CH2Data.xlsx"

    column_num = 3
    #file_location = "C:/Users/Tariq/Documents/impfile/EXCELL STUFF.xlsx"
    workbook = xlrd.open_workbook(file_location)

    # I guess sheets are named starting from 0
    sheet = workbook.sheet_by_index(0)
   #sheet = workbook.sheet_by_index("name")    # Uncomment and give the sheets name to use this

    l = []
    # The first row is considered row 0 so this could get confusing
    for i in range(0, 400):
        a = sheet.cell_value(i, column_num)   # Gets the value in the cell. (row number, column number)
        if a == "":
            break
        else:
            
            #print("row " + str(i+1), sheet.cell_value(i,7))
            l.append((i+1, sheet.cell_value(i,column_num)))

    return l

#(row, column value_left (column 7 in excel), value_right (column 8 in excel)   

def local_min(l):
    new_l = []
    temp = []
    repeat = False
    temp = []

    print(l[0])
    
    # Makes assunption beginning few is not repeated
    for i in range(1, len(l)-1):
        if l[i][1] < l[i+1][1] and l[i][1] < l[i-1][1]:
            new_l.append(l[i])

        #Checks if point next to i is the same
        elif l[i][1] == l[i+1][1]:
            if repeat == False:
                point = l[i-1][1]
                temp.append(l[i])
                repeat = True

            elif repeat == True:
                temp.append(l[i])   #If not repeated, add pointer to i-1.
                # If repeated, add element to the repeated list
        
        elif l[i][1] != l[i+1][1] and repeat: #Checks when repeat is done
            if point > temp[0][1] and l[i+1][1] > temp[0][1]: #Putting first cond >= to check if list begins with repeats
                new_l.extend(temp)
                new_l.append(l[i])
                repeat = False
                temp = []

            else:
                repeat = False
                temp = []
                

    # If it ends in repeats, add them to the list
    if len(temp) > 0:
        new_l.extend(temp)

        
    return new_l



def local_max(l):
    new_l = []
    temp = []
    repeat = False
    temp = []

    # Makes assunption beginning few is not repeated
    for i in range(0, 400):  #len(l)-1
        if l[i][2] > l[i+1][2] and l[i][2] > l[i-1][2]:
            new_l.append(l[i])

        #Checks if point next to i is the same
        elif l[i][2] == l[i+1][2]:
            if repeat == False:
                point = l[i-1][2]
                temp.append(l[i])
                repeat = True

            elif repeat == True:
                temp.append(l[i])   #If not repeated, add pointer to i-1.
                # If repeated, add element to the repeated list
        
        elif l[i][2] != l[i+1][2] and repeat: #Checks when repeat is done
            if point < temp[0][2] and l[i+1][2] < temp[0][2]: #Putting first cond >= to check if list begins with repeats
                new_l.extend(temp)
                new_l.append(l[i])
                repeat = False
                temp = []

            else:
                repeat = False
                temp = []
                

    # If it ends in repeats, add them to the list
##    if len(temp) > 0:
##        new_l.extend(temp)

        
    return new_l

## NOTE: Hasn't been tested for cases where all numbers are the same


l = extract()
#print(l)
min_points = local_min(l)
#max_points = local_max(l)

print(min_points)
#print(max_points)


