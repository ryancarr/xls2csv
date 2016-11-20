# Program     : xls2csv.py
# Author      : Ryan Carr
# Created     : 03/07/15
# Description : Program opens an Excel worksheet and converts it
#               to a comma delimited file. Menu driven system
#               requires .xls or .xlsx files to be in the same
#               folder as the .py file.

from xlrd import open_workbook
from os import getcwd, listdir
from os.path import isfile

def askquestion(question):
    ''' askquestion(question) -> string

        Waits for user input based on question parameter
        Questions should be yes or no format
        Returns 'y' for yes and 'n' for no
    '''
    while True:
        inp = raw_input(question)
        if inp.lower().startswith('y'):
            result = 'y'
        elif inp.lower().startswith('n'):
            result = 'n'
        else:
            continue
        break
    return result

def converttocsv(fname, delimiter=','):
    ''' converttocsv(fname, delimiter=',')

        Opens and converts a Excel xls file to comma delimited csv
    '''
    intro = 'Which worksheet do you want to open?'
    line = ""

    # Check if the file exists. If it does alert user and ask if they want
    # to overwrite it. If no prompt for another filename.
    while True:
        ofname = raw_input('Enter a filname to save as (include .csv): ')
        if isfile(ofname):
            question = '{} already exists. Overwrite? (y/n): '.format(ofname)
            answer = askquestion(question)
            if answer == 'y':
                break
            elif answer == 'n':
                continue
        else:
            break

    # Open xls file with Unicode support, close program if failure
    try:
        workbook = open_workbook(fname, encoding_override='cp1252')
    except:
        print "{} is not a valid Excel file. Closing program.".format(fname)
        raise SystemExit

    # Generates a list of all the sheet names in the workbook
    question = []
    for sheet in workbook.sheet_names():
        question.append(sheet)

    # Ask user which sheet they want to convert
    sheet = displaymenu(intro, question)
    worksheet = workbook.sheet_by_name(sheet)

    # Create a new file and prepare it for writing
    with open(ofname, "w") as fh:

        # Nested loops allow us to navigate the 2D array of cells
        for curRow in range(worksheet.nrows):
            for curCol in range(worksheet.ncols):
                # Generate a single line of output with a delimiter
                line += unicode(worksheet.cell(curRow, curCol).value) + delimiter
            # Encode line to UTF-8 then write to file
            fh.write((line + "\n").encode("utf8"))
            line = ""

    print "Successfully completed converting {0} to {1}.".format(fname, ofname)

def displaymenu(intro, lst):
    ''' displaymenu(lst) -> object

        Displays a menu and allows the user to choose an option.
        Returns their choice as an object.
    '''
    counter = 0     
     
    while True:
        print intro
          
        for var in lst:
            print '{0}) {1}'.format(str(counter), var)
            counter += 1
        try:
            inp = int(raw_input('Enter a number: '))
        except:
            print "Invalid input"
            continue
        if 0 <= inp and inp < len(lst): break
        else:
            print 'Selection isn\'t on the list. Try again.\n'
            counter = 0
    return lst[inp]

def getfiles(fileext='.xls'):
    ''' getfiles(fileext) -> list_of_strings

        Returns a list of all files with fileext extension
    '''
    files = listdir(getcwd())
    files.sort(key=str.lower)

    # This loop reduces the list to files with a specific
    # file extension
    templst = []
    for var in files:
        if fileext not in var[var.find('.'):]: continue
        templst.append(var)
    
    return templst

def main():
    cont = True
    while cont:
        files = getfiles()
        if not files:
            print 'There aren\'t any .xls files in ' + getcwd()
            break
        intro = 'This program converts files from .xls or .xlsx to .csv'
        fname = displaymenu(intro, files)
        converttocsv(fname)

        answer = askquestion("Would you like to convert another file? (y/n): ")
        if answer == 'y': continue
        elif answer == 'n': cont = False


if __name__ == '__main__':
    main()
