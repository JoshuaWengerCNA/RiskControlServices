import pandas as pd
from normalize import getLocList
# from normalize import getNormalData

def reconcileData(df1, df2):
    locations = getLocList(df2)
    dataFrame = df1[~df1['Location ID'].isin(locations)]
    return dataFrame

# TESTS:

# firstDF = getNormalData(pd.read_excel('.\Source File Samples\Covered or Endorsed by CNA.xlsx'))
# secondDF = getNormalData(pd.read_excel('.\Source File Samples\Covered in JO.xlsx'))

# addCoverageList = reconcileData(firstDF, secondDF)
# pullCoverageList = reconcileData(secondDF, firstDF)

# with pd.ExcelWriter('ExcelTest.xlsx') as writer:
#     addCoverageList.to_excel(writer, sheet_name = 'Review for Adding Coverage', index = False)
#     pullCoverageList.to_excel(writer, sheet_name = 'Review for Removing Coverage', index = False)
