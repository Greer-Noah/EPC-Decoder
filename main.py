from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import os
import pandas as pd
# import tkinterDnD
from pyepc import SGTIN
from pyepc.exceptions import DecodingError
import csv
from csv import reader
import xlsxwriter
import openpyxl

# path = "/Users/noahgreer/Desktop"
# os.chdir(pathE

# epcList = []
# itemfiles = []

# folder_name = "WeeklySupplierTrackingReport{1}".format(store_num_input, user_date_input)
# folder_path = os.path.join(os.path.expanduser("~"), "Desktop", folder_name)
# os.mkdir(folder_path)
count = 0
"""
select_txt() Function:
    The select_txt function prompts the user to select a group of Input text files.
    The function then compiles the EPCs listed in the text files into a list, of which duplicate EPCs are removed.
    This list of EPCs is then written into a single column in a new Input Excel File (.xlsx), 
        named "Store[Store #]CC[MMDDYYYY]", which is then saved on to the user's Desktop.
    The Input file location is then passed into the DecodeCycleCount(file_location).
"""

def select_txt():
    epcList = []
    # --------------Prompts User for Input files------------------------------------------------------------------
    filenames = filedialog.askopenfilenames(initialdir = "/", title = "Select Input Text Files",
                                           filetypes = (("txt files", "*.txt"), ("all files", "*.*")))

    # --------------Reads Input Files-----------------------------------------------------------------------------
    for filename in filenames:
        f = open(filename, "r")
        lines = f.readlines()
        for x in lines:
            epcList.append(x.split('\n')[0])
        f.close()

    # --------------Deletes EPC duplicates------------------------------------------------------------------------------
    epcList_noDupe = [*set(epcList)]

    # --------------Exports Input input file----------------------------------------------------------------------
    df1 = pd.DataFrame(epcList_noDupe, columns=['EPCs'])
    cc_file_name = "Input_File.xlsx"
    path1 = os.path.join(os.path.expanduser("~"), "Desktop", cc_file_name)
    str(path1)
    global count
    count += 1
    if count > 1:
        path2 = path1.split('.')
        path2.insert(-1, ' ({0}).'.format(count-1))  # Adds '_output' to filename
        path1 = ''.join(path2)
    writer = pd.ExcelWriter(path1, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='EPCs', startrow=0, startcol=0, index=False)
    writer.save()

    # --------------Starts decoding process-----------------------------------------------------------------------------
    print("Preparing to Decode...")
    # print("Path1: {0}".format(path1))
    DecodeCycleCount(path1)


def select_csv():
    epcList = []
    global count
    count += 1
    # --------------Prompts User for Input files------------------------------------------------------------------
    filenames = filedialog.askopenfilenames(initialdir = "/", title = "Select Input CSV Files",
                                           filetypes = (("csv files", "*.csv"), ("all files", "*.*")))

    # --------------Reads Input Files-----------------------------------------------------------------------------
    for filename in filenames:
        df = pd.read_csv(filename)
        epcList = df.values.tolist()

        # with open(filename, 'rb') as file:
        #     print(chardet.detect(file.read()))

    print(epcList)
    # --------------Deletes EPC duplicates------------------------------------------------------------------------------
    res = list(map(''.join, epcList))
    epcList_noDupe = [*set(res)]

    # --------------Exports Input input file----------------------------------------------------------------------
    df1 = pd.DataFrame(epcList_noDupe, columns=['EPCs'])
    cc_file_name = "Input_File.xlsx"
    path1 = os.path.join(os.path.expanduser("~"), "Desktop", cc_file_name)
    str(path1)
    if count > 1:
        path2 = path1.split('.')
        path2.insert(-1, ' ({0}).'.format(count-1))  # Adds '_output' to filename
        path1 = ''.join(path2)
    writer = pd.ExcelWriter(path1, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='EPCs', startrow=0, startcol=0, index=False)
    writer.save()


    # --------------Starts decoding process-----------------------------------------------------------------------------
    print("Preparing to Decode...")
    # print("Path1: {0}".format(path1))
    DecodeCycleCount(path1)


def select_xlsx():
    epcList = []
    global count
    count += 1
    # --------------Prompts User for Input files------------------------------------------------------------------
    filenames = filedialog.askopenfilenames(initialdir = "/", title = "Select Input Excel Files",
                                           filetypes = (("xlsx files", "*.xlsx"), ("all files", "*.*")))

    # --------------Reads Input Files-----------------------------------------------------------------------------
    for filename in filenames:
        df = pd.read_excel(filename)
        epcList = df.values.tolist()

    print(epcList)
    # --------------Deletes EPC duplicates------------------------------------------------------------------------------
    res = list(map(''.join, epcList))
    epcList_noDupe = [*set(res)]

    # --------------Exports Input input file----------------------------------------------------------------------
    df1 = pd.DataFrame(epcList_noDupe, columns=['EPCs'])
    cc_file_name = "Input_File.xlsx"
    path1 = os.path.join(os.path.expanduser("~"), "Desktop", cc_file_name)
    str(path1)
    if count > 1:
        path2 = path1.split('.')
        path2.insert(-1, ' ({0}).'.format(count-1))  # Adds '_output' to filename
        path1 = ''.join(path2)
    writer = pd.ExcelWriter(path1, engine='xlsxwriter')
    df1.to_excel(writer, sheet_name='EPCs', startrow=0, startcol=0, index=False)
    writer.save()


    # --------------Starts decoding process-----------------------------------------------------------------------------
    print("Preparing to Decode...")
    # print("Path1: {0}".format(path1))
    DecodeCycleCount(path1)

"""
DecodeCycleCount(file_location) Function:
    The DecodeCycleCount(file_location) function takes in the location of the input Input file previously created
        and reads the file into a list.
    The function then attempts to decode each EPC in the list.
    If the EPC, or SGTIN, successfully decodes into a UPC, or GTIN, then that GTIN is stored in a UPC list.
    Every EPC that could not be properly decoded into a UPC is then recorded in an Error EPC list.
    This EPC list has a corresponding Error Message which is recorded into the Error UPC list.
    The function then creates a Input Output Excel file (.xlsx) with three corresponding sheets:
        1. A list of all of the inputted EPCs and their corresponding UPCs. (Sheet Name: 'Unique EPCs, Dupe UPCs')
        2. A list of all of the non-duplicated UPCs. (Sheet Name: 'Unique EPCs, Unique UPCs')
        3. A list of all of the incorrectly decoded EPCs and their corresponding error messages. (Sheet Name: 'Errors')
    This output file is saved onto the user's Desktop.
    Lastly, the function passes the 'duplicates-included' UPC list into the sql_connect_populate(upcList) function.
"""


def DecodeCycleCount(file_location):

    # --------------Reads in the input Input----------------------------------------------------------------------
    df2 = pd.read_excel(file_location)

    epcList1 = []

    columns = df2.columns.tolist()

    # --------------Creates EPC list using input CC file----------------------------------------------------------------
    for _, i in df2.iterrows():
        for c in columns:
            epcList1.append(i[c])

    upcList = []
    errorEPCs = []
    errorUPCs = []

    # --------------Decoding Process------------------------------------------------------------------------------------
    print("Decoding...")
    for x in epcList1:
        try:
            upcList.append(SGTIN.decode(x).gtin) # Actual decode command
        except DecodingError as e: # Handles decoding errors
            errorEPCs.append(x)
            errorUPCs.append(e)
        except TypeError as t: # Handles decoding errors
            errorEPCs.append(x)
            errorUPCs.append(t)

    # --------------Removes Error EPCs from EPC list--------------------------------------------------------------------
    for epc in errorEPCs:
        if epc in epcList1:
            epcList1.remove(epc)

    epcList2 = []
    upcList2 = []

    # --------------Creates distinct UPC list---------------------------------------------------------------------------
    for i in range(len(upcList)):
        if upcList[i] not in upcList2:
            upcList2.append(upcList[i])
            epcList2.append(epcList1[i])

    # --------------Deletes leading 0s from UPCs------------------------------------------------------------------------
    for y in range(len(upcList)):
        upcList[y] = upcList[y].lstrip('0')

    # --------------Deletes leading 0s from UPCs------------------------------------------------------------------------
    for z in range(len(upcList2)):
        upcList2[z] = upcList2[z].lstrip('0')

    # --------------Formats and prepares lists as Pandas DataFrames to be exported to .xlsx-----------------------------
    df3 = pd.DataFrame(epcList1, columns=['EPCs'])
    df4 = pd.DataFrame(upcList, columns=['UPCs'])
    # df5 = pd.DataFrame(epcList2, columns=['EPCs'])  # EPC list on 'Unique EPC, Unique UPC' sheet
    df6 = pd.DataFrame(upcList2, columns=['UPCs'])
    df7 = pd.DataFrame(errorEPCs,columns=['EPCs'])
    df8 = pd.DataFrame(errorUPCs, columns=['UPCs'])

    # --------------Exports Output file with Duplicate UPCs, Unique UPCs, and Errors------------------------
    cc_file_name = "Output_File.xlsx"
    path10 = os.path.join(os.path.expanduser("~"), "Desktop", cc_file_name) # Saves on Desktop
    str(path10)
    global count
    if count > 1:
        path2 = path10.split('.')
        path2.insert(-1, ' ({0}).'.format(count-1))  # Adds file count to filename
        path10 = ''.join(path2)
    writer = pd.ExcelWriter(path10, engine='xlsxwriter')

    print("Creating Output File...")
    df3.to_excel(writer, sheet_name='Unique EPCs, Dupe UPCs', startrow=0, startcol=0, index=False)
    # df5.to_excel(writer, sheet_name='Unique EPCs, Unique UPCs', startrow=0, startcol=0, index=False)
    df4.to_excel(writer, sheet_name='Unique EPCs, Dupe UPCs', startrow=0, startcol=1, index=False)
    df6.to_excel(writer, sheet_name='Unique EPCs, Unique UPCs', startrow=0, startcol=0, index=False)
    df7.to_excel(writer, sheet_name='Errors', startrow=0, startcol=0, index=False)
    df8.to_excel(writer, sheet_name='Errors', startrow=0, startcol=1, index=False)

    writer.save()


ws = tk.Tk()
def UserInterfaceCreation():
    ws.title("EPC Decoder")
    w = 800
    h = 650
    sw = ws.winfo_screenwidth()
    sh = ws.winfo_screenheight()
    x = (sw/2) - (w/2)
    y = (sh/2) - (h/2)
    ws.geometry('%dx%d+%d+%d' % (w, h, x, y))
    # ws.wm_attributes('-transparent', True)
    ws.config(bg="#CCE1F2")

    quit_button = ttk.Button(ws, text="Quit", width=8, command=Close)
    # button3.place(x=300, y=150)
    quit_button.place(relx=.5, rely=.9, anchor=CENTER)


    # ---------------------------------------------------------------------------------------------
    # # Create an Entry widget to accept User Input
    # entry = Entry(ws, width=40)
    # entry.focus_set()
    # #entry.place(x=300, y=100)
    # entry.place(relx=.5, rely=.35, anchor=CENTER)
    # ---------------------------------------------------------------------------------------------

    # Create a Button to enter the select_txt() function and select the Input .txt files
    button1 = ttk.Button(ws, text="Select Input Files (.txt)", command=select_txt)
    button1.place(relx=.5, rely=.4, anchor=CENTER)

    button2 = ttk.Button(ws, text="Select Input Files (.csv)", command=select_csv)
    button2.place(relx=.5, rely=.5, anchor=CENTER)

    button3 = ttk.Button(ws, text="Select Input Files (.xlsx)", command=select_xlsx)
    button3.place(relx=.5, rely=.6, anchor=CENTER)


def Close():
    ws.destroy()
    SystemExit(0)


if __name__ == '__main__':
    UserInterfaceCreation()
    ws.mainloop()
    # Ends the program
    exit(0)