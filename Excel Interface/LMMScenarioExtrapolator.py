'Prophet interpolated ZCs'
'--------------------------'
import pandas as pd
import numpy as np
from openpyxl import *

# Import esg table
dataLMMPlus = pd.read_pickle('YE19 Scenarios.pkl')

maturities = [1, 2, 3, 4, 5, 10, 15, 20, 25, 30, 35, 40]
indexFilter = ['ESG.Economies.EUR.NominalZCBP(Govt, '+str(i)+', 3)' for i in maturities]


# Interpolator used in Prophet
def ProphetInterpolator(maturities, zeroCouponBonds, maturityToExtrapolate):
    # Should be of equal length to the length of ZC Bonds
    Nb = len(maturities)
    
    # Forward Rate proxy
    FRate = np.log(zeroCouponBonds[Nb-1]/zeroCouponBonds[Nb-2])/ (maturities[Nb-1] - maturities[Nb-2])
    
    i = Nb
    while (maturityToExtrapolate < maturities[i-1]):
        i = i - 1
    
    # Interpolator by choice of maturity   
    if (maturityToExtrapolate == maturities[i -1]):
        result = zeroCouponBonds[i - 1]
    
    # Main interpolator
    elif(maturityToExtrapolate < maturities[Nb -1]):
        x = ((np.exp(-(maturityToExtrapolate - maturities[i - 1]) * FRate) -
            np.exp( -(maturities[i] - maturities[i - 1]) * FRate))/
            (1 - np.exp(-(maturities[i] - maturities[i - 1]) * FRate)))
             
        result = zeroCouponBonds[i-1] * x + (1 - x) * zeroCouponBonds[i]
    
    # Extrapolator
    else:
        result = zeroCouponBonds[Nb -1] * np.exp(-(maturityToExtrapolate - maturities[Nb-1]))* FRate
        
    return (result)

# Implementation of the interpolator on the ZC Table
finalDataSet = []
for i in range(1, 3001):
    subdata = dataLMMPlus.loc[i].set_index('Parameter')
    zcScenarios  = subdata.loc[indexFilter].to_numpy().transpose()
    
    subdataExtrapolated = []
    for item in zcScenarios:
        subdataExtrapolated.append([ProphetInterpolator(maturities, item, j) for j in np.arange(1, 31)])
    
    finalDataSet.append(pd.DataFrame(np.transpose(subdataExtrapolated)))
  
# Concatenate dataframes
finalDataFrame = pd.concat(finalDataSet)
finalDataFrame.to_excel('LMMPlus Interpolated Rates (Python).xlsx')   

'Find the deflateur'
'------------------------------'
deflateurs = np.cumprod(finalDataFrame.loc[0], axis = 1)
deflateurs = deflateurs.drop(deflateurs.columns[-1], axis = 1).to_numpy()

averageZC = []
avgZCDiscounted =[]
# Should be the average of the first row
for i in np.unique(finalDataFrame.index):
    averageZC.append(np.mean(finalDataFrame.loc[i], axis = 0))
    avgZCDiscounted.append(np.mean(np.multiply(finalDataFrame.loc[i].drop(0, axis = 1).to_numpy(), 
                                               deflateurs), axis = 0))
    
# Find the average ZC in TTM
avgZC = pd.DataFrame(averageZC, 
                     index = np.arange(1, len(averageZC)+1), 
                     columns = np.arange(1, 51)).transpose()

avgZCDiscounted = pd.DataFrame(avgZCDiscounted, 
                     index = np.arange(1, len(averageZC)+1), 
                     columns = np.arange(1, 51)).transpose()

# Obtain the average defalteur
avgDeflateur = np.mean(np.cumprod(finalDataFrame.loc[0], axis = 1), axis = 0)


' Write to Excel'
'-----------------------'
book = load_workbook('Martingale Test.xlsx')
writer = pd.ExcelWriter('Martingale Test.xlsx', engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
   
avgZC.to_excel(writer,
               "Average ZC", 
               startcol = 2, # Column F
               startrow = 1, # Line 59
               index = True, 
               header = True)

avgZCDiscounted.to_excel(writer,
               "Martingale Test", 
               startcol = 6, # Column F
               startrow = 1, # Line 59
               index = True, 
               header = True)

pd.DataFrame(avgDeflateur).to_excel(writer,
               "Martingale Test", 
               startcol = 2, # Column F
               startrow = 1, # Line 59
               index = False, 
               header = False)

writer.close()





































