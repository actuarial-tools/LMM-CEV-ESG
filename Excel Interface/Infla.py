# -*- coding: utf-8 -*-
"""
Created on Wed Jul 29 08:25:56 2020

@author: s082360
"""

"MODIFY THE SCENARIO TABLES"
'---------------------------------'
import pandas as pd
import numpy as np
import os

# Set Working Directory
loc = '//cerata/DGCU/Risques/16. Actuarial Function/05. Economic Scenario Générator/00 - DD LMM CEV/02 - Travaux/Excel Interface'
os.chdir(loc)

# Import both scenario tables  
tableLMMPlus = pd.read_pickle("YE19 Scenarios.pkl").set_index('Parameter', append=True)
tableLMMCEV = pd.read_csv("Results\DD LFM CEV YE19 Scenarios.csv", index_col = [0, 1])

# For each scenario modify the inflation table 
for i in np.unique([elem[0] for elem in tableLMMCEV.index]):
    subDataLMMPlus = tableLMMPlus.loc[i]
    subDataCEV = tableLMMCEV.loc[i]
    
    inflaLMMPlus = (subDataLMMPlus.loc['ESG.Economies.EUR.RealZCBP(Govt, 1, 3)']/
                    subDataLMMPlus.loc['ESG.Economies.EUR.NominalZCBP(Govt, 1, 3)']) -1
    
    inflaCEV = (subDataCEV.loc['ESG.Economies.EUR.RealZCBP(Govt, 1, 3)']/
                    subDataCEV.loc['ESG.Economies.EUR.NominalZCBP(Govt, 1, 3)']) -1
    
    tableLMMPlus.loc[i].loc['ESG.Economies.EUR.InflationRates.Inflation.Index'] = np.append([1], np.cumprod(1 + inflaLMMPlus[:50]))
    tableLMMCEV.loc[i].loc['ESG.Economies.EUR.InflationRates.Inflation.Index'] = np.append([1], np.cumprod(1 + inflaCEV)[:50])

tableLMMPlus.to_csv('Results\LMMPlus Inflation Adjusted Scenarios.csv')
tableLMMCEV.to_csv('Results\LMMCEV Inflation Adjusted Scenarios.csv')
