from openpyxl import load_workbook #Import workbook
from openpyxl import Workbook #Create and import workbook
import multiprocessing
from multiprocessing import Pool, Queue
import time
from copy import copy, deepcopy
from operator import itemgetter


###########################################################################################
############# Retrieve excel file to read

# set file path to retrieve file
#filepath="C:\\Users\\AJ Suchovsky\\Desktop\\Multiprocess examples\\Absence_Roster.xlsx"
filepath="C:\\Users\\suchovaj\\Desktop\\Parallel Python\\Absence_Roster.xlsx"
#filepath="C:\\Users\\suchovaj\\Desktop\\Parallel Python\\Large_Roster_Test.xlsx"

# load demo.xlsx 
wb=load_workbook(filepath,data_only = True)

# select demo.xlsx
sheet=wb.active

# get max row count
max_row=sheet.max_row

# get max column count
max_column=sheet.max_column


############################################################################################
########### Appending another excel sheet with sorted values ###############################
######## This is to open that new workbook which will be used latter #######################

filepath2="C:\\Users\\suchovaj\\Desktop\\Parallel Python\\Sorted_Roster_Blank.xlsx"

wb2=load_workbook(filepath2)

sheet2=wb2.active

###########################################################################################

def print_read(d,k,lock):
    lock.acquire()
    alist = []
    
    for i in d:
        #time.sleep(0.5)
    # iterate over all columns
        blist = []
        for j in range(1,max_column+1):
          # get particular cell value    
            cell_obj=sheet.cell(row=i,column=j)
          # print cell value     
          #list.append(cell_obj.value)
            blist.append(cell_obj.value)
        alist.append(blist)
     # print new line
    print(alist)
    k.put(alist)
    lock.release()

def find_ms_lvl(do_num, use_queue, ms1,ms2,ms3,ms4,lock):
    lock.acquire()
    cadet = []
    if use_queue.empty() == False:
        for i in do_num:
            cadet.extend(use_queue.get())
            if cadet[2] == 1:
                ms1.put(cadet)
            elif cadet[2] == 2:
                ms2.put(cadet)
            elif cadet[2] == 3:
                ms3.put(cadet)
            else:
                ms4.put(cadet)
        lock.release()
    else:
        lock.release()

def full_run():
    full_list = []
    full_list2 = []
    ms1_list = []
    ms2_list = []
    ms3_list = []
    ms4_list = []
    main_list = []
    
    ms1_queue = multiprocessing.Queue()
    ms2_queue = multiprocessing.Queue()
    ms3_queue = multiprocessing.Queue()
    ms4_queue = multiprocessing.Queue()

    full_queue = multiprocessing.Queue()

    lock = multiprocessing.Lock()

    quart = int(max_row/4)
    quart1 = multiprocessing.Process(target=print_read, args=(range(2,quart),full_queue,lock))
    quart2 = multiprocessing.Process(target=print_read, args=(range(quart,quart*2),full_queue,lock))
    quart3 = multiprocessing.Process(target=print_read, args=(range(quart*2,quart*3),full_queue,lock))
    quart4 = multiprocessing.Process(target=print_read, args=(range(quart*3,max_row+1),full_queue,lock))
    ms_lvl = multiprocessing.Process(target=find_ms_lvl, args=(range(2,max_row),full_queue,ms1_queue,ms2_queue,ms3_queue,ms4_queue,lock))

    quart1.start()
    quart2.start()
    quart3.start()
    quart4.start()
    ms_lvl.start()
    
    quart1.join()
    quart2.join()
    quart3.join()
    quart4.join()
    ms_lvl.join()
    
    while not ms1_queue.empty():
        print(ms1_queue.get())

    while not ms2_queue.empty():
        print(ms2_queue.get())

    while not ms3_queue.empty():
        print(ms3_queue.get())

    while not ms4_queue.empty():
        print(ms4_queue.get())

###########################################################################################
if __name__ == '__main__':
    starttime = time.time()

    full_run()

    wb2.save(filepath2)
    print('That took {} seconds'.format(time.time() - starttime))
