import pandas as pd

policyDocPath='.\Source File Samples\Covered or Endorsed by CNA.xlsx'
joDocPath='.\Source File Samples\Covered in JO.xlsx'

# Read in data to data frames
polInputDF = pd.read_excel(policyDocPath, sheet_name='Locations') # Column named: 'JO Loc ID'
joInputDF = pd.read_excel(joDocPath) # Column named: 'Location ID'

# Funtcion getLocList() returns a list of location IDs
def getLocList(dataFrame):
    return dataFrame['Location ID'].tolist()
    # locList = dataFrame['Location ID'].values.tolist()
    # return locList

def fixLocDetailHeaders(dataFrame):
    if('Juris. #' in dataFrame.columns):
        dataFrame.rename(columns={'Juris. #':'Juris#'}, inplace=True)
    if('Cert Expire Date' in dataFrame.columns):
        dataFrame.rename(columns={'Cert Expire Date':'Cert Expires'}, inplace=True)
    if('Next Inspect Due' in dataFrame.columns):
        dataFrame.rename(columns={'Next Inspect Due':'Due Date'}, inplace=True)
    if('Juris. #' in dataFrame.columns):
        dataFrame.rename(columns={'Juris. #':'Juris#'}, inplace=True)
    if('Policy' in dataFrame.columns):
        dataFrame[['Policy Name', 'Policy Number']] = dataFrame['Policy'].str.split('(', 2, expand = True)
        dataFrame['Policy Number'] = dataFrame['Policy Number'].str.replace(')','')
    return dataFrame
    # Fix headers
    # Split 'Policy' column to 'Policy Name' and 'Policy Number'
	# Original: | <policy name> (<policy number>) |
	# Revised:  | <policy name> | <policy number> |

# Create normalized DF for JO data
def normalizeJoInput(dataFrameIn):
    dataFrame = dataFrameIn
    print(dataFrame['Location ID'].dtype)
    if(dataFrame['Location ID'].dtype != int):
        print('Not int')
        dataFrame['Location ID'] = dataFrame['Location ID'].astype(int)
    dataFrame = dataFrame.sort_values(by = 'Location ID')
    dataFrame['Object Count'] = dataFrame.groupby('Location ID')['Location ID'].transform('count')
    dataFrame['Due Date'] = pd.to_datetime(dataFrame['Due Date'], errors='coerce')
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
    # dataFrame = dataFrame[dataFrame['Active Objects'] > 0] # **Moved operation to compareData()
    # dataFrame[['Account Number', 'Building Values', 'New/Renewal', 'SIC-Occupancy', 'LOB', 'Policy Start']]\
    #     = dataFrame['Client Reference Field'].str.split('|', 6, expand = True)
    # dataFrame[['SIC', 'Occupancy']] = dataFrame['SIC-Occupancy'].str.split(' ', 1, expand = True)
    try:
        pipeCount = dataFrame.iloc[2]['Client Reference Field'].count('|')
    except:
        pipeCount = 0
        tempDF = pd.DataFrame()
        tempDF['CRF-1'] = dataFrame['Client Reference Field']
    print('Pipe Count: ',pipeCount)
    if(pipeCount>0):
        tempDF = dataFrame['Client Reference Field'].str.split('|', expand = True)
        tempDF.columns = ['CRF-'+str(field+1) for field in tempDF.columns]
        # dataFrame[['CRF-1', 'CRF-2', 'CRF-3', 'CRF-4', 'CRF-5', 'CRF-6']]\
        #     = dataFrame['Client Reference Field'].str.split('|', pipeCount+1, expand = True)
    dataFrame.rename(columns={'JO Loc ID': 'Location ID', 'Location Name': 'Policyholder Name', 'Active Objects': 'Object Count',\
                               'Next Due Date (Active Objs Only)':'Due Date'}, inplace=True)
    # normDF = dataFrame[['Account Number', 'Location ID', 'Policyholder Name', 'JO Location Name', 'Address', 'JO Address', 'City', 'JO City', 'State', \
    #                     'JO State', 'Zip', 'JO Zip', 'Object Count', 'Due Date', 'Building Values', 'New/Renewal', 'SIC', 'Occupancy',\
    #                     'LOB', 'Policy Start']] # Modified to add JO Location Name and JO Address Columns
    normDF = dataFrame[['Location ID', 'Policyholder Name', 'JO Location Name', 'Address', 'JO Address', 'City', 'JO City', 'State', \
                    'JO State', 'Zip', 'JO Zip', 'Object Count', 'Due Date']]
                    # Modified to add JO Location Name and JO Address Columns
    normDF = pd.concat([normDF, tempDF], axis=1)
    # normDF.drop_duplicates(subset = 'Location ID', inplace = True) # **Moved operation to compareData()
    return normDF

def getNormalData(dataFrame):
    colNames = list(dataFrame.columns.values)
    if 'JO Loc ID' in colNames:
        return normalizeMatchInput(dataFrame)
    # elif 'Location ID' in colNames:
    #     return normalizeJoInput(dataFrame)
    else:
        return normalizeJoInput(dataFrame)
        # print('error')
        # return
