from datetime import datetime
from tkinter import filedialog as fSelector
from getpass import getuser

def getOutputDir():
    stamp = datetime.now().strftime('%y%m%d')
    directory = 'C:/Users/' + getuser() + '/Downloads/Reconciled Locations ' + str(stamp) + '.xlsx'
    return directory

def getFile(title,dir='',extParams=['any']):
    # Gets name of user to open correct directory
    userName = getuser()

    # Builds directory location string
    dirString = 'C:\\Users\\' + userName + '\\' + dir + '\\'

    # Builds list of file types to allow for selection
    extList=fileExtensions(extParams)

    selection = fSelector.askopenfilename(
        title=title,
        initialdir=dirString,
        filetypes=extList
    )
    return selection

# fileExtenions() takes a list of extensions and returns a list of filetypes strings
def fileExtensions(ext):
    # Dictionary of file type strings keyed by extension abbreviation
    extDict={
        'xlsx': ('Excel file', '.xls, .xlsx'),
        'txt': ('text files', '.txt'),
        'any': ('All files', '*.*')
    }
    feList=[]
    # If single filetype, returns single string
    if type(ext) != list:
        if ext not in extDict:
            return extDict['any']
        else:
            extension=extDict[ext]
            feList.append(extension)
            return feList

    # Builds list of filetypes
    for extension in ext:
        extString=extDict[str(extension)]
        feList.append(extString)
    return feList
