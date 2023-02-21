import tkinter
import pandas as pd
import numpy as np
import os
from datetime import datetime
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
import tkinter.filedialog
import tkinter as tk
import openpyxl
import shutil


'''
STATUS: 
    -COMPLETE; BUT ONLY WORKS ONE BY ONE TAB
    
GOAL:
    TO COMBINE MULTIPLE TABS INTO ONE FILE. 

STEPS:
    1 -     USER WILL PASS A FILE WITH MULTI TABS THEY WANT TO COMBINE AND A NEW EXCEL FILE WITH THE HADERS ALREADY ADDED [must be csv]
    2 -     USER WILL INDICATE WHICH TAB THEY WANT TO ADD TO THE BLANK FILE
    3 -     USER WILL THEN MAP THE COLUMNS BETWEEN THE TWO FILES
    4 -     SCRIPT WILL GET COLUMN DATA FROM BOTH, COMBINE AND ADD THEM TO AN EMPTY DF
    5 -     ONCE ALL COLUMNS ARE ITERATED, SCRIPT WILL SAVE THE NEW DATAFRAME TO THE NAME OF THE BLANK FILE USER PASSED ON
    
TO ADD:
    0 - AS IT STANDS, ALL TABS MUST BE ADDING THE SAME COLUMNS. 
        after it iterates through the first tab, itll reset the headers to whatever the uesr selected to import in the first tab
        i.e column A and B mapping for first tab, the combined csv will have only those 2
        when moving on to the next tab, other columns wont exist anymore
    **** MUST CHANGE THE LEVEL OF ITER THE DF IS MAINTAINED
    1 - SHOW THE LIST OF TABS ADDED IN UI #2 
    2 - CREATE A COPY OF BOTH FILES FOR INCASE PURPOSES IN A FOLDER IN THE CURRENT WORKING DIR
    3 - CHECK FOR FIRST ROW; MAKE SURE ITS THE HEADERS WE NEED
    4 - OPTION TO SKIP A TAB THAT WAS ADDED BUT BASED ON SOMETHING ELSE, THEY DONT WANT TO ADD
    5 - BACK BUTTON; OR ALT  
        A - OPTION TO MODIFY TABS LIST
        B - OPTION TO MODIFY HEADERS LIST 
    6 - IF ONE ENTRY BOX BLANK THEN ERR_MSG -
    7 - PRINT 
        A - TABS LIST AS IT IS UPDATED
        B - COLUMN MAPPING

****
when adding multiple columns from mutli tabs, it keeps erasing after the second one
if i move the newdf to a level above the current, it only keeps the first tone
'''

################ FUNCTIONS ################
## get the string from the entry box; for submit button
def get_data():

    # set variables from this function to global so they can be used elsewhere
    global datafilepath
    global bslfilepath
    global datafile_df
    global bsl_df
    global datafile_sheetList
    global bsl_columnList

    # get value [filepath] from entry box's
    datafilepath = ent1.get() ## xlsx
    bslfilepath = ent2.get() ## csv
    wb = openpyxl.load_workbook(datafilepath) ## create wb obj for the excel file

    ## if they dont already exist, create copy of bsl file
    directory = os.path.dirname(bslfilepath)
    base_filename = os.path.basename(bslfilepath)
    newfile = directory + '\\[COPY] Baseline Ready data_file.csv'

    if os.path.exists(newfile) == True:
        pass
    elif os.path.exists(newfile) == False:
        shutil.copy(fn_merge,newfile)


    # show user the files that have been passed
    print(f"Files chosen are:\n"
          f"- Data File: {datafilepath}\n"
          f"- Baseline File: {bslfilepath}")

    # create data frame files
    datafile_df = pd.read_excel(datafilepath)
    bsl_df = pd.read_csv(bslfilepath)

    # get column headers into a list [to pass to OptionsMenu (dropdown)]
    datafile_sheetList = wb.sheetnames ## for the excel file
    bsl_columnList =  list(bsl_df.columns.values) ## for the csv bsl file

    # # checks
    # print(datafile_columnList)
    # print(bsl_columnList)

    # close the UI in order to use data and open next one
    root.destroy()


## browse files function for datafile file
def browsefunc1():
    # open file explorer + file types
    filename =tkinter.filedialog.askopenfilename(filetypes=(("All files","*.*"),("xls files","*.xls"),("xlsx files","*.xlsx")))
    # clear the entry box before inserting file path [to prevent errors in filepath]
    ent1.delete(0,tk.END)
    # add this to the entry box
    ent1.insert(tk.END, filename)


## browse files function for bsl file
def browsefunc2():
    # open file explorer + file types
    filename =tkinter.filedialog.askopenfilename(filetypes=(("csv files","*.csv"),("All files","*.*")))
    # clear the entry box before inserting file path [to prevent errors in filepath]
    ent2.delete(0,tk.END)
    # add this to the entry box
    ent2.insert(tk.END, filename) # add this to the entry box


## grab tabs user wants to combine
def grab():
    # get tabname
    tab = variable.get()
    # msg to user
    print(f"Adding {tab}")
    # add tabname into list
    tabs_list.append(tab)



## EXE THE QA TOOL SCRIPT ##
def execute():
    # close the UI
    root.destroy()
    # show user the tabs
    print('\nTabs to combine:')
    for tab in tabs_list:
        print('\t',tab)


## EXE THE QA TOOL SCRIPT ##
def execute_script():
    root.destroy()


################ FUNCTIONS END ################


######################################### UI #1 #########################################

################ UI DETAILS ################

## initiate tkinter UI and settings
root = tk.Tk()
root.title("Combine tabs into One") # title
screenwidth = root.winfo_screenwidth() #dynamic screenwidth based on computer
screenheight = root.winfo_screenheight() #dynamic screenheight based on computer
height = 200
width = 350
# dynamic UI sizeing
alignstr = '%dx%d+%d+%d' % ((width + screenwidth/10),(height +screenheight/10),(screenwidth - width) / 2, (screenheight - height) / 2) ## dynamic sizing for the screen
# set size of UI
root.geometry("400x400")
## allow resizing from the user
root.resizable(width=True, height=True)
## App Name
AppLabel = tk.Label(root, text='Combine tabs into One')



################ BUTTON 1; PATH TO DATA FILE ################

## label for the base file
label= tk.Label(root, text="Data file here: ", font=('Verdana 13'))
label.place(x=50,y=0)


# browse button
b1=tk.Button(root,text="BROWSE",font=40,command=browsefunc1,bd="4")
b1.place(x=10,y=30)


# entry box where the filepath will appear
ent1=tk.Entry(root,font=40)
ent1.place(x=100,y=35)



################ BUTTON 2; PATH TO FILE DATA WILL BE IMPORTED TO; BSL_FILE ################
## label for the comp file
label= tk.Label(root, text="Destination file here:", font=('Verdana 13'))
label.place(x = 50,y = 80)

# browse button
b2=tk.Button(root,text="BROWSE",font=40,command=browsefunc2,bd="4")
b2.place(x = 10,y = 110)

# entry box 2 where the filepath will appear
ent2=tk.Entry(root,font = 40)
ent2.place(x = 100,y = 115)


################ BUTTON 3; CONFIRMS FILEPATHS AND MOVES TO NEXT STEP ################

## button to confirm the selected files are the one the user wants
getbutton = tk.Button(root, text= "SUBMIT",bg='blue',command=get_data,bd="8",fg="white")
getbutton.place(x = 150,y = 150)

## print message to user in case of ayn error
msg_to_user = tk.Label(root, text="Please submit two files", font=('Verdana 13'))
msg_to_user.place(x=100,y=200)


root.mainloop()

#########################################  UI #1 #########################################



######################################### UI #2 #########################################

## start the second UI to get the column headers
root = tk.Tk()
root.title("Selecting Tabs")
root.geometry("400x400")
root.resizable(width=True, height=True) ## allow resizing from the user
AppLabel = tk.Label(root, text='Selecting Tabs')

tabs_list = []

# Drop down for selecting the tabs the user wants to add
# dropdown menu for 1; basefile
variable = tk.StringVar(root)
variable.set(datafile_sheetList[0])
datafile_DropDownMenu = tk.OptionMenu(root,variable,*datafile_sheetList)
datafile_DropDownMenu.place(x=40,y=40)

# label to indicate which files headers these are
datafile_dropdown_label = tk.Label(root, text="Base File:", font=('Verdana 13'))
datafile_dropdown_label.place(x = 40,y = 10)


## BUTTON TO GRAB THE DATA AND ADD TO LIST
grab = tk.Button(root,text='add',command=grab)
grab.place(x = 300,y = 40)


## execute button
execute = tk.Button(root, text= "execute",command=execute,bg='dark green',bd="8",fg="white")
execute.place(x = 130,y = 90)


root.mainloop()

######################################### UI #2 END #########################################

######################################### UI #3 start #########################################
## grab column pairs and put into list
def grab():
    print(f"Adding {variableb.get()} | {variableb.get()}")
    datafile_column = variablea.get()
    bsl_column = variableb.get()
    datafile_column_mapping.append(datafile_column)
    bsl_column_mapping.append(bsl_column)

print(tabs_list)
for tabs in tabs_list:
    # create empty dataframe to store the data
    newdf = pd.DataFrame()
    print(f"Processing {tabs}")

    ## create dataframe for excel sheet; will be opening up different tabs
    df = pd.read_excel(datafilepath,tabs)

    ## get list of columns names from excel tab
    datafile_columnsList = list(df.columns.values)

    ## start the second UI to get the mapping
    root = tk.Tk()
    root.title("Mapping Columns")
    root.geometry("400x400")
    root.resizable(width=True, height=True) ## allow resizing from the user
    AppLabel = tk.Label(root, text='Mapping Columns')

## mapping the columns between the data file and the destination bsl file

    bsl_column_mapping = []
    datafile_column_mapping = []

# drop down
    # dropdown menu for 1; data file columns
    variablea = tk.StringVar(root)
    variablea.set(datafile_columnsList[0])
    basefile_DropDownMenu = tk.OptionMenu(root,variablea,*datafile_columnsList)
    basefile_DropDownMenu.place(x=40,y=40)

    # label to indicate which files headers these are
    compfile_dropdown_label = tk.Label(root, text="Data File", font=('Verdana 13'))
    compfile_dropdown_label.place(x = 40,y = 10)

# drop down
    # dropdown menu for 2; bsl file columns
    variableb = tk.StringVar(root)
    variableb.set(bsl_columnList[0])
    basefile_DropDownMenu = tk.OptionMenu(root,variableb,*bsl_columnList)
    basefile_DropDownMenu.place(x=40,y=140)

    # label to indicate which files headers these are
    compfile_dropdown_label = tk.Label(root, text="Baseline Ready File", font=('Verdana 13'))
    compfile_dropdown_label.place(x = 40,y = 110)

# add button
    ## grab button to add the two columns selected by user and put them into a list
    grab = tk.Button(root,text='add',command=grab)
    grab.place(x = 300,y = 40)


## execute button
    execute = tk.Button(root, text= "execute",command=execute_script,bg='dark green',bd="8",fg="white")
    execute.place(x = 40,y = 190)

# main tkitner loop
    root.mainloop()

## excel and csv dtaframe of two files; dataframe and bsl ready
    datafile_df = pd.read_excel(datafilepath,tabs)
    bsl_df = pd.read_csv(bslfilepath)

    print(bsl_column_mapping)
    print(datafile_column_mapping)

    print(datafile_df)
# loop through the # of columns the user has mapped
    print('\tCombing columns from both files')
    for columnindex in range(0,len(bsl_column_mapping)):

        # variable for the column names
        columnName = datafile_column_mapping[columnindex] ## data file column name
        columnName_bsl = bsl_column_mapping[columnindex] ## bsl file column name

        # show user the column mapping as they are iterated through
        print('\t '+ columnName+ '  |  ' +columnName_bsl)

        # get column data and turn into a series from datafile
        series_1 = pd.Series(datafile_df[columnName])

        # get column data and turn into a series from bslfile
        series_2 = pd.Series(bsl_df[columnName_bsl])

        # combine both series into one and give the column the name of the data file [the correct header bsl will take]
        series_3 = pd.concat([series_1,series_2],axis=0)
        print(series_3)
        # remove the blanks/blank rows in the file;
        # vendors usually drag the dropdown so while they look empty, pandas detects them as cells with values
        series_3.dropna(axis=0,
                        how='all',
                        inplace=True)

        # add combined series to the df
        newdf[columnName_bsl] = series_3

        #
    newdf.to_csv(bslfilepath,index=False)

