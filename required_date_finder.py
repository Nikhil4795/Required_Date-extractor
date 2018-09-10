"""
    Written @ Nikhil

    If you have any questions or suggestions regarding this script,
    feel free to contact me via nikhil.ss4795@gmail.com

    Report bugs to @ nikhil.ss4795@gmail.com
    

    Description :
    
    Finds earliest or latest date when given a stream of dates.
    1. When given only one date it returns the same date.
    2. When mentioned date seperator is not present, then it remains same.
    

    Example :
            Input  : 11/21/2001 | 4/9/2001 | 9/14/2002  (Month-Day-year format)
            Output : 4/9/2001 (Earliest)
            Output : 9/14/2002 (Latest)

            Input : 11/21/2001
            Output : 11/21/2001 (Earliest)
            Output : 11/21/2001 (Latest)

    Input Sheet Constraints :
            1. Must be either csv or xlsx. (recommended : csv)

    output sheet :
            1. The output sheet contains all the columns from the input sheet and
               the results of the earliest or latest date is resulted on the last column.

            2. If the input sheet is xlsx and the dates column also contains single date in a cell then,
                a. The output sheet may contain a flaw, the rows which doesn't have
                       multiple dates may result in some other format (number format).
                b. So, to convert those single date cell into date format, from number
                       format follow the steps mentioned below.
                c. Replace if any punctuation from the result column to blanks. (Usually the punctuations are : "[","]","'")
                d. Then select the whole column and right click it, then move to "format cells" section,
                       and then select date from the left section and then you can choose the
                       date format which you required and then hit ok.
                e. That's it, now your result column is in date format.

            3. If the input sheet is csv, then it works fine. No need to do any external formatting.

    How to run :
            1. Keep only two files in the folder. one should be script file and the other is input sheet.
            2. Hit run
            3. Mention the date format of the input sheet from the available date formats.
            4. Mention the date which you required, i.e.., either earliest date or latest date.
            5. Mention the seperator used in your input sheet to seperate dates with quotes and without spaces.
            6. Mention the column number starting from 0, in which the dates are present.
                
            
"""

import csv
import xlrd, xlsxwriter
import datetime
import os
import sys
import time

file_names = os.listdir('.')

if len(file_names) == 2 :
    pass
else:
    print("wrong number of files in the folder : ")
    print("     Please keep only two files in folder (input_sheet and script file)")
    sys.exit()

input_type = ""
input_name = ""

for i in range(0,len(file_names)):
    if file_names[i][-4:] == 'xlsx':
        input_name = file_names[i][:-5]
        input_type = "xlsx"
    if file_names[i][-3:] == 'csv':
        input_name = file_names[i][:-4]
        input_type = "csv"

if input_type == 'csv' or input_type == 'xlsx':
    print("Input file type found : ",input_type)
else:
    print("Error in file format : ")
    print("     Sheet must be either in csv or xlsx ")
    sys.exit()

print("Date Formats : ")
print("     1. Month-Day-Year")
print("     2. Day-Month-Year")
print("     3. Year-Month-Day")

date_format = input("Enter any choice from the mentioned date formats : ")

avaiable_date_format = ["%m/%d/%Y","%d/%m/%Y","%Y/%m/%d"]

if date_format == 1 or date_format == 2 or date_format == 3 :
    pass
else:
    print("\nWrong Date format mentioned")
    print("     Please enter either 1 or 2 or 3 according to sheet dates format")
    sys.exit()

print("\nEnter 1 to extract earliest date from multiple dates")
print("Enter 2 to extract latest date from multiple dates")
date_choice = input("Please Enter your choice : ")

if date_choice == 1 or date_choice == 2:
    pass
else:
    print("\nWrong choice mentioned")
    print("     Please enter either 1 or 2")
    sys.exit()



try:
    print("\n\nPlease mention the seperator used to seperate your dates")
    print("Example : 11/19/2001 | 5/21/2008 | 8/15/2003")
    print("""     Here '|' is seperating the multiple dates""")
    date_seperator = input("\nEnter your date seperator with quotes : ")
except:
    print("\nError : Wrong Format of mentioning the seperator")
    print("        Please make sure that you've entered the seperator within quotes")
    sys.exit()


print("\n\nMention in which column the dates are present")
date_col = input("Enter the colum number starting from 0 : ")


    
start_time = time.time()

def sort_date(input_given):
    latestdate=""
    now=datetime.datetime.now().date()
    now_date=str(now)
    temp_date = 1
    for i in range(0,len(input_given)):
        temp = input_given[i].replace(" ","")
        try:
            new_date=datetime.datetime.strptime(temp,avaiable_date_format[date_format - 1]).strftime("%Y-%m-%d")
        except:
            print("\nError : Mentioned date format is incorrect, please recheck it again")
            sys.exit()
        if date_choice == 1:
            if(new_date < now_date):
                latestdate=temp
                now_date=new_date
        if date_choice == 2:
            if(new_date > temp_date):
                latestdate=temp
                temp_date=new_date
            
    return latestdate

if input_type == 'csv':
    
    datain = csv.reader(open(input_name+".csv","rU"))
    dataout = open(input_name+"_results.csv","wb")
    datawrite = csv.writer(dataout)
    header = []
    
    for index,each in enumerate(datain):
        if index == 0:
            for j in range(0,len(each)):
                header.append(each[j])
            header.append("Required Date")
            datawrite.writerow(header)
        else:
            row_data = []
            for i in range(0,len(each)):
                row_data.append(each[i])
            try:
                temp = each[date_col].split(date_seperator)
            except:
                print("\nError : Mentioned dates column number is Incorrect. Please recheck it again.")
                sys.exit()
            final_temp = []
            for i in range(0,len(temp)):
                final_temp.append(temp[i].strip())

            if len(final_temp) > 1:
                latest = sort_date(final_temp)
            else:
                latest = final_temp

            row_data.append(latest)
            datawrite.writerow(row_data)

if input_type == 'xlsx':
    
    datain = xlrd.open_workbook(input_name+".xlsx").sheet_by_index(0)
    dataout = xlsxwriter.Workbook(input_name+"_results.xlsx")
    datawrite = dataout.add_worksheet()
    
    header = []
    
    
    for row in range(0, datain.nrows):
        if row == 0 :
            for j in range(0, datain.ncols):
                header.append(datain.cell_value(row,j))

            header.append("Required Date")

            for col in range(0,len(header)):
                datawrite.write(row,col,header[col])
        else:
            row_data = []
            for j in range(0, datain.ncols):
                row_data.append(str(datain.cell_value(row,j)))

            try:
                temp = str(datain.cell_value(row,date_col)).split(date_seperator)
            except:
                print("\nError : Mentioned dates column number is Incorrect. Please recheck it again.")
                sys.exit()
            
            final_temp = []
            for i in range(0,len(temp)):
                final_temp.append(temp[i].strip())

            if len(final_temp) > 1:
                latest = sort_date(final_temp)
            else:
                latest = str(final_temp)

            row_data.append(str(latest))

            for col in range(0,len(row_data)):
                datawrite.write(row,col,str(row_data[col]))
   

print("Task Completed")
print("     Took around : %s seconds " % (time.time() - start_time))
dataout.close()

"""

    Written @ Nikhil

    If you have any questions or suggestions regarding this script,
    feel free to contact me via nikhil.ss4795@gmail.com

    report bugs to @ nikhil.ss4795@gmail.com

"""
