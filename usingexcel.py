# In the command prompt, write __pip install thing____

import xlrd


def extract():
    """
    Obtains all the information needed to find the local max and mins
    """
    file_location = "C:/Users/Tariq/Downloads/CH2Data.xlsx" #CHANGE: Location should point to the place where the file is kept

    column_num = 3 #CHANGE: This is the column number of the values that would go in the x axis of a graph if you were to plot it - 1 since we
                   #start with 0 when counting in CS

    workbook = xlrd.open_workbook(file_location)

    # Sheet indexes start from 0
    sheet = workbook.sheet_by_index(0)
   #sheet = workbook.sheet_by_index("name")    # Uncomment this and comment the line above it to find the sheet by it's name if it has a name

    l = []
    # The first row is considered row 0 instead of row 1, just like before with the columns
    for row in range(0, 400):
        inside_cell = sheet.cell_value(row, column_num)   # Gets the value in the cell given by (row number, column number)
        if inside_cell == "":
            break
        else:
            
            #print("row " + str(i+1), sheet.cell_value(i,7))
            l.append((row+1, sheet.cell_value(row,column_num)))

    return l


def local_min(l):
    new_l = []
    temp = []
    repeat = False
    temp = []


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
            if point > temp[0][1] and l[i+1][1] > temp[0][1]:
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
            if point < temp[0][2] and l[i+1][2] < temp[0][2]: 
                new_l.extend(temp)
                new_l.append(l[i])
                repeat = False
                temp = []

            else:
                repeat = False
                temp = []
                

        
    return new_l


data = extract()
min_points = local_min(data)
#max_points = local_max(data)          # Comment out either max_points or min_points, whichever one you don't want right now. After this, comment out
                                       # the corresponding print statement
print(min_points)
#print(max_points)


