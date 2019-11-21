# import load_workbook
from openpyxl import load_workbook
import multiprocessing
from multiprocessing import Pool
import time
# set file path
#filepath="C:\\Users\\AJ Suchovsky\\Desktop\\Multiprocess examples\\Absence_Roster.xlsx"
filepath="C:\\Users\\suchovaj\\Desktop\\Parallel Python\\Absence_Roster.xlsx"
# load demo.xlsx 
wb=load_workbook(filepath,data_only = True)
# select demo.xlsx
sheet=wb.active
# get max row count
max_row=sheet.max_row
# get max column count
max_column=sheet.max_column
# iterate over all cells 
# iterate over all rows


list = [] #go back an name this main list
list_for_1 = []
list_for_2 = []
list_for_3 = []
list_for_4 = []


def print_read1(k):
    alist = []
    #for i in range(1,int(max_row/4)):
     
    # iterate over all columns
    for j in range(1,max_column+1):
          # get particular cell value    
        cell_obj=sheet.cell(row=k,column=j)
          # print cell value     
          #list.append(cell_obj.value)
        alist.append(cell_obj.value)  
     # print new line
    return alist

def send_ms1_list(k):
    list_ms1 = []
    #for i in list:
    #for n in k:
    #for n in list[k]:
    if list[k][:][2] == 1:
        for j in range(1,max_column+1):
                # get particular cell value    
            cell_obj=sheet.cell(row=k,column=j)
                # print cell value     
                #list.append(cell_obj.value)
            list_ms1.append(cell_obj.value) 
    return list_ms1

def send_ms2_list():
    list_ms1 = []
    return list_ms2

def send_ms3_list():
    list_ms1 = []
    return list_ms3

def send_ms4_list():
    list_ms1 = []
    return list_ms4

def sort_to_lists():
    for i in list:
        for j in i:
            if list[i][j][2] == 1:
                list_for_1.append(list[i][j])
            elif list[i][j][2] == 2:
                list_for_2.append(list[i][j])
            elif list[i][j][2] == 3:
                list_for_3.append(list[i][j])
            elif list[i][j][2] == 4:
                list_for_4.append(list[i][j])

if __name__ == '__main__':
    p1 = Pool(processes=10)
    p2 = Pool(processes=10)
    p3 = Pool(processes=10)
    p4 = Pool(processes=10)

    data1 = p1.map(print_read1, [i for i in range(2,int(max_row/4))])
    data2 = p2.map(print_read1, [i for i in range(int(max_row/4),int(max_row/4)*2)])
    data3 = p3.map(print_read1, [i for i in range(int(max_row/4)*2,int(max_row/4)*3)])
    data4 = p4.map(print_read1, [i for i in range(int(max_row/4)*3,max_row+1)])

    p1.close()
    p2.close()
    p3.close()
    p4.close()

    list.append(data1)
    list.append(data2)
    list.append(data3)
    list.append(data4)

    print(list)
    #p1 = multiprocessing.Process(target=sort_to_lists)

    p5 = Pool(processes=10)
    p6 = Pool(processes=10)
    p7 = Pool(processes=10)
    p8 = Pool(processes=10)

    ms1 = p5.map(send_ms1_list, (i for i in enumerate(list[:])))
    ms2 = p6.map(send_ms2_list, (i for i in enumerate(list[:])))
    ms3 = p7.map(send_ms3_list, (i for i in enumerate(list[:])))
    ms4 = p8.map(send_ms4_list, (i for i in enumerate(list[:])))

    p5.close()
    p6.close()
    p7.close()
    p8.close()

    list_for_1.append(ms1)
    list_for_2.append(ms2)
    list_for_3.append(ms3)
    list_for_4.append(ms4)
    
    
    print(list_for_1)
    print(list_for_2)
    print(list_for_3)
    print(list_for_4)

    #print(list[0][2])

