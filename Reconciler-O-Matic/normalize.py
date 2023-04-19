import pandas as pd

policyDocPath='.\Source File Samples\Covered or Endorsed by CNA.xlsx'
joDocPath='.\Source File Samples\Covered in JO.xlsx'

# Read in data to data frames
polInputDF = pd.read_excel(policyDocPath, sheet_name='Locations') # Column named: 'JO Loc ID'
joInputDF = pd.read_excel(joDocPath) # Column named: 'Location ID'

# Funtcion getLocList() returns a list of location IDs
def getLocList(dataFrame):
    locList = dataFrame['Location ID'].tolist()
    return locList


# Create normalized DF for JO data
def normalizeJoInput(dataFrameIn):
    dataFrame = dataFrameIn
    dataFrame = dataFrame.sort_values(by = 'Location ID')
    dataFrame['Object Count'] = dataFrame.groupby('Location ID')['Location ID'].transform('count')
    dataFrame['Due Date'] = pd.to_datetime(dataFrame['Due Date'])
    dataFrame['Due Date'] = dataFrame['Due Date'].dt.date
    byDate = dataFrame.sort_values(by = 'Due Date')
    dataFrame = byDate.groupby('Location ID').head()
    dataFrame.drop_duplicates(subset = 'Location ID', inplace = True)
    dataFrame = dataFrame.sort_values(by = 'Location ID')
    dataFrame.rename(columns={'Location Name': 'JO Location Name'}, inplace=True) # Added
    normalDF = dataFrame[['Location ID','JO Location Name', 'Address', 'City', 'State', 'Zip', 'Object Count', 'Due Date', 'Policy Number']]
    return normalDF

def normalizeMatchInput(dataFrameIn):
    dataFrame = dataFrameIn
    dataFrame = dataFrame[dataFrame['Active Objects'] > 0]
    dataFrame[['Account Number', 'Building Values', 'New/Renewal', 'SIC-Occupancy', 'LOB', 'Policy Start']]\
        = dataFrame['Client Reference Field'].str.split('    ', 6, expand = True)
    dataFrame[['SIC', 'Occupancy']] = dataFrame['SIC-Occupancy'].str.split(' ', 1, expand = True)
    dataFrame.rename(columns={'JO Loc ID': 'Location ID', 'Location Name': 'Policyholder Name', 'Active Objects': 'Object Count',\
                               'Next Due Date (Active Objs Only)':'Due Date'}, inplace=True)
    normDF = dataFrame[['Account Number', 'Location ID', 'Policyholder Name', 'JO Location Name', 'Address', 'JO Address', 'City', 'JO City', 'State', \
                        'JO State', 'Zip', 'JO Zip', 'Object Count', 'Due Date', 'Building Values', 'New/Renewal', 'SIC', 'Occupancy',\
                        'LOB', 'Policy Start']] # Modified to add JO Location Name and JO Address Columns
    normDF.drop_duplicates(subset = 'Location ID', inplace = True)
    return normDF

def getNormalData(dataFrame):
    colNames = list(dataFrame.columns.values)
    if 'JO Loc ID' in colNames:
        return normalizeMatchInput(dataFrame)
    elif 'Location ID' in colNames:
        return normalizeJoInput(dataFrame)
    else:
        print('error')
        return
