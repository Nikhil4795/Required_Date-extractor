# Required_Date-extractor

Extracts Earliest date or Latest date when given a stream of dates.

Written @ Nikhil

    If you have any questions or suggestions regarding this script,
    feel free to contact me via nikhil.ss4795@gmail.com

    report bugs to @ nikhil.ss4795@gmail.com
    

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
    
    
    Requirements :
        1. python 2 or python 3.
        2. Modules required in python :
            csv, xlrd, xlsxwriter, os, sys, time, datetime.
            (To install modules : pip install module_name)
    
    Written @ Nikhil

    If you have any questions or suggestions regarding this script,
    feel free to contact me via nikhil.ss4795@gmail.com

    report bugs to @ nikhil.ss4795@gmail.com
