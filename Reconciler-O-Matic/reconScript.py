# Script to automatically compare values of SOV and JOL data sets an provide matches where found.
# Takes .xlsx SOV and JOL Location Report as input and generates a .xlsx with results
# Author: Joshua Wenger - EB Risk Control Specialist - joshua.wenger@cna.com

# The pyinstaller command below builds the executable file for distribution:
# Pyinstaller.exe --onefile --name Reconciler-O-Matic --windowed --icon=favicon.ico reconScript.py

# Imports
import pandas as pd
import os
from tkinter import *
from getFile import getFile, getOutputDir
from normalize import getNormalData
from compareData import reconcileData
from stat import S_IREAD, S_IRGRP, S_IROTH # Can be used to set read only



# Select spreadsheets for comparison
def grabSheet(sheetName):
    # global policyDocPath, joDocPath, policyDF, policyInputCount, joDF, joInputCount
    global policyDF, joDF
    if sheetName == 'policy':
        policyDocPath = getFile('Open Policy File', 'Downloads', 'xlsx')
        polPath.config(text = os.path.split(policyDocPath)[1], bg='grey')
        policyDF = pd.read_excel(policyDocPath, sheet_name='Locations')
        polStats.config(text = 'Policy data contains ' + f"{len(policyDF):,}" + ' line items.')
    elif sheetName == 'jo':
        joDocPath = getFile('Open JOL Data File', 'Downloads', 'xlsx')
        joPath.config(text = os.path.split(joDocPath)[1], bg='grey')
        joDF = pd.read_excel(joDocPath)
        joStats.config(text = 'Policy data contains ' + f"{len(joDF):,}" + ' line items.')

def compareData():
    # Set output file name
    outFileName = getOutputDir()
    print(outFileName)

    # Get normalized and deduped data frames
    normalAddDF = getNormalData(policyDF)
    normalRemoveDF = getNormalData(joDF)

    # Reconcile the SOV data with the JO data
    reconAddDF = reconcileData(normalAddDF, normalRemoveDF)
    reconRemoveDF = reconcileData(normalRemoveDF, normalAddDF)

    # # Get some stats
    addCoverageCount = len(reconAddDF)
    removeCoverageCount = len(reconRemoveDF)

    # Write Normalized and deduped data frames to excel file
    with pd.ExcelWriter(outFileName) as writer:
        reconAddDF.to_excel(writer, sheet_name = 'Review for Adding Coverage', index = False)
        reconRemoveDF.to_excel(writer, sheet_name = 'Review for Removing Coverage', index = False)

    # Set file to read only
    # os.chmod(outFileName, S_IREAD|S_IRGRP|S_IROTH)

    # Launch newly created excel spreadshet in MS Excel
    os.startfile(outFileName)

# Create GUI window with Tkinter
main = Tk()
main.title(' Data Reconciler-O-Matic 5000 v0.1')
main.geometry('403x550')
main.config(bg="white", padx=10, pady=20)
main.iconbitmap("favicon.ico")

topLabel = Label(main, pady=25,font=('Verdana', 12, 'bold'), bg='red', fg='white', text='Choose the SOV-based matching template:')
topLabel.grid(row=0, columnspan=3)

# Interface
getPolSheet = Button(main, text = 'Open Policy File', command = lambda : grabSheet('policy'))
getPolSheet.grid(row=1, column=1, pady=15)

# Show Policy spreadsheet path
polPath = Label(main, text = '', bg = 'white')
polPath.grid(row=2, columnspan=3)

# Show Policy spreadsheet stats
polStats = Label(main, text = '', bg = 'white')
polStats.grid(row=3, columnspan=3, pady=(5,15))

joSheetLabel = Label(main, pady=25,font=('Verdana', 12, 'bold'), bg='red', fg='white', text='Choose the JOL-based data spreadsheet:')
joSheetLabel.grid(row=4, columnspan=3)

getJoSheet = Button(main, text = 'Open JOL Data File', command = lambda : grabSheet('jo'))
getJoSheet.grid(row=5, column=1, pady=15)

# Show JO spreadsheet path
joPath = Label(main, text = '', bg = 'white')
joPath.grid(row=6, columnspan=3)

# Show JO spreadsheet stats
joStats = Label(main, text = '', bg = 'white')
joStats.grid(row=7, columnspan=3, pady=(5,15))

compareLabel = Label(main, pady=25,font=('Verdana', 12, 'bold'), bg='red', fg='white', \
                     text='Click ''Reconcile Data'' to compare the data\nsets and view'\
                          ' the reconciled spreadsheet')
compareLabel.grid(row=8, columnspan=3)

compareAndOpen = Button(main, text = 'Reconcile Data', command = lambda : compareData())
compareAndOpen.grid(row=9, column=1, pady=15)

main.mainloop()
