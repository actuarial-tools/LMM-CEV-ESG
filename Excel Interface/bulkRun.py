'NOTE'
'---------------'
'* All computation is done within the esgaviva module'
'* This allows for modifying the module variables and functions'

# Manage operating systems
import os
import numpy as np
import pandas as pd
import PySimpleGUI as sg
import xlsxwriter

# Set Working Directory
loc = '//cerata/DGCU/Risques/16. Actuarial Function/05. Economic Scenario Générator/00 - DD LMM CEV/02 - Travaux/Excel Interface'
os.chdir(loc)

# Initialize Curve Setup
import esgaviva # Very important because in this module we will modify some values in the CalibSetup module
import imp

#wb = esgaviva.xlw.Book.caller()
wb = esgaviva.xlw.Book('IR Nominal ESG Tool.xlsm')

sg.theme('DarkGrey2')

       

'1)  MANAGEMENT OF SCENARIOS GENERATED'
'=============================================='

def singleRun(inputs, scenariosAddress, auditTrailAddress, maxMaturity, maxProjYear):
    'Run Calibration'
    '-----------------------'
    resultsCalibration = esgaviva.haganCalibratorPython(inputs[:6], eta = inputs[6], delta = inputs[7])
    
    'Run Projection'
    '-----------------------'
    results = esgaviva.BHfullSimulator(esgaviva.forwardCurve, 
                    resultsCalibration[0],
                    resultsCalibration[1],
                    resultsCalibration[2],
                    resultsCalibration[3],
                    resultsCalibration[4],
                    resultsCalibration[5], inputs[6], esgaviva.betas, inputs[7])
    
    'MODIFYING THE CAPECO SCENARIO TABLE '
    '----------------------------------------------'
    'Convert Scenarios from Triangle to Rectangular shape'
    # Save Distributions in Prophet Format
    simulatedCurves = results[0]
    deflateurs = results[1]
    
    # Sortir les prix Zero Coupon  
    zCoupons = esgaviva.copy.deepcopy(simulatedCurves)
    
    # Transform Forwards to ZC
    for i in range(len(simulatedCurves)):
        for j in range(len(simulatedCurves[0])):
            zCoupons[i][j] = np.cumprod(1/(1+simulatedCurves[i][j]))
    
    # Change to Time to Maturity - Select the diagonals
    ZCTimeToMaturity = esgaviva.copy.deepcopy(simulatedCurves)
    for i in range(len(zCoupons)):
        for j in range(maxMaturity): # Select the jth item from each vector ie the diagonals 
            ZCTimeToMaturity[i][j] = [x[j] for x in zCoupons[i][:(len(zCoupons[i][0])-j)]]  #Vector of length maturity + projection - j
            
    #Select the boundaries required (maxProjYear + maxMaturity)
    ZCdistributions = np.zeros((maxMaturity*len(ZCTimeToMaturity), maxProjYear))
    ratesdistributions = np.zeros((maxMaturity*len(ZCTimeToMaturity), maxProjYear))
    for i in range(len(ZCTimeToMaturity)):
        for j in range(maxMaturity):
            ZCdistributions[i*maxMaturity+j] = np.array(ZCTimeToMaturity[i][j][:maxProjYear])
            ratesdistributions[i*maxMaturity+j] = np.power(1/ZCdistributions[i*maxMaturity+j], 1/(j+1))-1
          
    'Import Scenario table'
    '-----------------------------'
    data2 = pd.read_pickle('YE19 Scenarios.pkl')        
    maturities = [1, 2, 3, 4, 5, 10, 15, 20, 25, 30, 35, 40]
    maturitiesFilter = [mat -1 for mat in maturities]
    
    # From data, set index to the parameter column
    data2 = data2.set_index('Parameter', append = True)

    # Find all the indices we will need for each scenario
    indexProphet = [list(np.array(maturities)+i*maxMaturity -1) for i in range(len(simulatedCurves))]
    indexProphet = list(esgaviva.itertools.chain.from_iterable(indexProphet))
    
    zcToAdd = list(esgaviva.zeroCouponCurve[maturitiesFilter])*len(simulatedCurves)
    ratesToAdd = list(np.power(1/esgaviva.zeroCouponCurve[maturitiesFilter], 
                               1/np.array(maturities))-1)*len(simulatedCurves)
    
    # Filter our data and add the current ZC and rate values
    finalZCdistributions = [np.append(zc,curve) for zc, curve in list(zip(zcToAdd, ZCdistributions[indexProphet]))]
    finalRatesdistributions = [np.append(rate,curve) for rate, curve in list(zip(ratesToAdd, ratesdistributions[indexProphet]))]
    
    # filter only the ZC and rates scenarios
    zcFilters = [str('ESG.Economies.EUR.NominalZCBP(Govt, '+ str(i)+', 3)') for i in maturities]  
    rateFilters = [str('ESG.Economies.EUR.NominalSpotRate(Govt, '+ str(i)+', 3)') for i in maturities]  
    
    for i in np.unique(data2.index.get_level_values(0)):
        subdata = data2.loc[i]

        'Modify Equity indices'
        '----------------------------'
        oldEquity1 = subdata.loc['ESG.Assets.EquityAssets.E_EUR.TotalReturnIndex'] 
        oldEquity2 = subdata.loc['ESG.Assets.EquityAssets.P_FRA.TotalReturnIndex'] 
        oldDF = np.cumprod(subdata.loc['ESG.Economies.EUR.NominalZCBP(Govt, 1, 3)'])
        newDF = np.cumprod(finalZCdistributions[len(maturities)*int(i-1)])
        subdata.loc['ESG.Economies.EUR.NominalYieldCurves.NominalYieldCurve.CashTotalReturnIndex'] = np.append([1], 
                                                                                                                   1/newDF[:-1])
        subdata.loc['ESG.Assets.EquityAssets.E_EUR.TotalReturnIndex'] = oldEquity1 * oldDF/newDF
        subdata.loc['ESG.Assets.EquityAssets.P_FRA.TotalReturnIndex'] = oldEquity2 * oldDF/newDF
        
        'Modify Nominal Rates'
        '----------------------------'
        filt = list(np.isin(data2.loc[i].index,zcFilters))
        filtrates = list(np.isin(data2.loc[i].index,rateFilters))
        subdata[filt] = finalZCdistributions[len(maturities)*int(i-1):len(maturities)*int(i)]
        subdata[filtrates] = finalRatesdistributions[len(maturities)*int(i-1):len(maturities)*int(i)]
        
        'Modify inflation'
        '---------------------'
        inflaCEV = (subdata.loc['ESG.Economies.EUR.RealZCBP(Govt, 1, 3)']/
                subdata.loc['ESG.Economies.EUR.NominalZCBP(Govt, 1, 3)']) -1
        subdata.loc['ESG.Economies.EUR.InflationRates.Inflation.Index'] = np.append([1], np.cumprod(1 + inflaCEV[:50]))
                 
        'Reassign Values to original data'
        data2.loc[i] = subdata.to_numpy()

    
    # Save Final Data as csv
    data2.to_csv(scenariosAddress, chunksize = 100)        
    
    
    'MARTINGALE TEST'
    '========================================'
    
    'FORWARDS TO ZERO COUPONS TO DISCOUNT FACTORS'
    '================================================'
    # Obtain Discount Factors for each scenario
    DF = [np.append(esgaviva.zeroCouponCurve[0],
                    esgaviva.zeroCouponCurve[0]*np.cumprod(df)) for df in deflateurs]
    
    # Average Deflateur
    AvgDeflateur = np.mean(DF, axis = 0) 
    
    # Obtain the Zero Coupon Bonds
    zCoupons = esgaviva.copy.deepcopy(simulatedCurves)
    zCouponsTilde = esgaviva.copy.deepcopy(simulatedCurves)
    
    # Calculate Zero Coupons and Deflated Zero Coupons
    ###################################################################################
    for i in range(len(simulatedCurves)):
        for j in range(len(simulatedCurves[0])):
            zCoupons[i][j] = np.cumprod(1/(1+simulatedCurves[i][j]))
            
            # Multiply each ZC by the discount factor to obtain Discount factors at each timestep
            zCouponsTilde[i][j] = zCoupons[i][j]*DF[i][j]
    
    # Calculate Average Deflated ZC
    ##################################################################################
    # Copy the triangle shape
    resMean = esgaviva.copy.deepcopy(zCoupons[0])
    
    # For calculation of the mean Deflated ZCoupons
    for i in range(len(resMean)):
        resMean[i] = [0]*len(resMean[i])
    
    # Calculate the mean Deflated ZCoupons    
    for i in range(len(zCouponsTilde)):
        for j in range(len(zCouponsTilde[0])):
            resMean[j] = np.nansum([resMean[j], zCouponsTilde[i][j]], axis = 0)
            
    # Final Mean len(simulatedCurves) = number of simulations        
    zCouponTildeAvg = [i/len(simulatedCurves) for i in resMean]   
    
    
    'CONVERT TRIANGLE TO RECTANGLE'
    '===================================='
    # Average deflated ZC (Take the First 30 maturities  for each year projected)
    rectangleZCTildeAvg = [zCouponTildeAvg[i][:30] for i in range(50)]
    
    # Error Calculation
    errorsMGTest = [abs((rectangleZCTildeAvg[i]/esgaviva.zeroCouponCurve[1+i:31+i]) -1)
                                for i in range(len(rectangleZCTildeAvg))]
        

    'MODEL PRICE CALCULATION'
    '------------------------------'
    # Calculate the swaption prices using Chi2 
    expiries = esgaviva.weightsLMMPlus['Expiry'].to_numpy(int) 
    tenors = esgaviva.weightsLMMPlus['Tenor'].to_numpy(int)

    strikes = esgaviva.forwardSwapRateVect(expiries, tenors)
            
    # Calculate chi square prices
    chiSquarePrices = esgaviva.calibrationBlackVect(1, strikes, 
                                                      strikes, 
                                                      expiries, 
                                                      tenors, 
                                                      resultsCalibration[0], 
                                                      resultsCalibration[1], 
                                                      resultsCalibration[2],
                                                      resultsCalibration[3],
                                                      resultsCalibration[4],
                                                      resultsCalibration[5], 
                                                      resultsCalibration[6],
                                                      resultsCalibration[7])  
    
    # Calculate bachelier prices
    normalPrices = esgaviva.normalPayerVect(1, strikes, strikes, expiries, tenors, esgaviva.volatilitiesLMMPlus['Value'])
    
    # Compare this to model prices
    errors = np.abs((normalPrices - chiSquarePrices)/normalPrices)
    cols = len(esgaviva.weightsLMMPlus.pivot(index = 'Expiry', columns = 'Tenor', values = 'Value'))
    rows = len(errors)/cols
    errors = np.matrix(errors.reshape((int(rows), cols))).T
    Chi2Prices = chiSquarePrices.reshape((int(rows), cols)).T
    normalRectPrices = normalPrices.reshape((int(rows), cols)).T
    
    'PROPHET PRICE CALCULATION'
    # Calculate the swaption prices using Chi2 
    expiriesProphet = np.tile(np.linspace(1, 80, 80), 10).astype(int)
    tenorsProphet = np.repeat(np.linspace(1, 10, 10), 80).astype(int)
    
    strikesProphet = esgaviva.forwardSwapRateVect(expiriesProphet, tenorsProphet)
    
    # Calculate chi square prices
    chiSquarePricesProphet = esgaviva.calibrationBlackVect(1, strikesProphet, 
                                                      strikesProphet, 
                                                      expiriesProphet, 
                                                      tenorsProphet, 
                                                      resultsCalibration[0], 
                                                      resultsCalibration[1], 
                                                      resultsCalibration[2],
                                                      resultsCalibration[3],
                                                      resultsCalibration[4],
                                                      resultsCalibration[5], 
                                                      0.7,
                                                      0.025) 
    
    bachelierModelVolatilitiesProphet = esgaviva.volNormalATMFunctionVect(expiriesProphet, 
                                                                          tenorsProphet,
                                                                          chiSquarePricesProphet)
    
    prophetTableAddress = auditTrailAddress[:len(auditTrailAddress) - 16]+'Vol Implicites(Prophet).xlsx'
    
    pd.DataFrame(bachelierModelVolatilitiesProphet.reshape((10, 80)), 
                 columns = np.linspace(1, 80, 80),
                 index = np.linspace(1, 10, 10)).transpose().to_excel(prophetTableAddress)
    

    'OBTAIN THE DISTRIBUTIONS'
    '========================================='
    # Select the maturities we want to analyze
    distributionMaturities = [1, 5, 10, 20] # Years to be considered
    quantiles = [0.5, 1, 5, 10, 25, 50, 75, 90, 95, 99, 99.5] # Quantiles that interest us
    ranges = [-1,-0.2, -0.10, -0.05, -0.025, 0, 0.025, 0.05, 0.1, 0.2, 0.3, 1] # These will be the ranges for the negative rates table # These will be the ranges for the negative rates table
    
    # To be used as the index in the creation of results table
    quantileIndex = np.append(quantiles, ['Mean', 
                          'Standard Deviation',
                          'Variance Coef'])
    finalZCDistributions = []
    finalRatesDistributions = []
    negativeRatesProportion = []
    ZCvariations = []
    
    # Obtain all the distributions for each maturity
    for maturity in distributionMaturities:
        'Modify ZC Variations'
        '---------------'
        distribs =  ZCdistributions[np.arange(maturity - 1, len(ZCdistributions), maxMaturity)]
        
        'Modify Rates'
        '----------------'
        initialRate = np.power(1/esgaviva.zeroCouponCurve[maturity - 1], 1/maturity) - 1
        rates = np.array([np.append([initialRate], elem) for elem in np.power(1/distribs, 1/maturity) -1])
        ratesStats = np.percentile(rates, quantiles, axis = 0).tolist()
        ratesStats.append(list(np.mean(rates, axis = 0))) # Add mean
        ratesStats.append(list(np.std(rates, axis = 0))) # Add sDev
        ratesStats.append(list(np.std(rates, axis = 0)/np.mean(rates, axis = 0))) # Add sDev scaled
        
        'Modify Rate Variation'
        '---------------------'
        ZCDiffVariance = np.diff(rates, axis = 1)
        zcStats = np.percentile(ZCDiffVariance, quantiles, axis = 0).tolist()
        zcStats.append(list(np.mean(ZCDiffVariance, axis = 0))) # Add mean
        zcStats.append(list(np.std(ZCDiffVariance, axis = 0))) # Add sDev
        zcStats.append(list(np.std(ZCDiffVariance, axis = 0)/np.mean(ZCDiffVariance, axis = 0))) # Add sDev scaled
        ZCvariations.append(zcStats)
        # Find rate variations
        finalZCDistributions.append(np.percentile(distribs, quantiles, axis = 0))
        finalRatesDistributions.append(ratesStats)
        
        negRates = []
        # Find the proportion of negative rates
        for i in range(len(rates.T)):
            negRates.append(np.histogram(rates.T[i], ranges)[0]/3000)
        
        negativeRatesProportion.append(np.transpose(negRates))
    
    
    # Create Excel File in the directory
    xlsxwriter.Workbook(auditTrailAddress).close()

    
    'SAVE ALL RESULTS TO AUDIT TRAIL'
    '====================================='
    # Save Surface & Weights
    book = esgaviva.load_workbook(auditTrailAddress)
    writer = pd.ExcelWriter(auditTrailAddress, engine = 'openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    
    pd.DataFrame(['Average Deflateur']).to_excel(writer, "Martingale Test", startcol = 3, startrow = 1, index = False, header = False)
    pd.DataFrame(['Input ZC Curve']).to_excel(writer, "Martingale Test", startcol = 2, startrow = 1, index = False, header = False)
    pd.DataFrame(['Martingale Test Error Table']).to_excel(writer, "Martingale Test", startcol = 6, startrow = 1, index = False, header = False)
    
    pd.DataFrame(esgaviva.zeroCouponCurve, index = np.arange(1, 151)).to_excel(writer, 
                                  "Martingale Test", 
                                  startcol = 1, # Column F
                                  startrow = 2, # Line 59
                                  index = True, 
                                  header = False)
        
    pd.DataFrame(AvgDeflateur[:50]).to_excel(writer, 'Martingale Test',
                                              startcol = 3, # Column F
                                                startrow = 2, # Line 38
                                                index = False, 
                                                header = False)

    pd.DataFrame(errorsMGTest, index = np.arange(1, 51)).to_excel(writer, 'Martingale Test',
                                              startcol = 6, # Column F
                                                startrow = 2, # Line 38
                                                index = True, 
                                                header = False)  
    
    'Calibration Report'
    '----------------------'
    pd.DataFrame(['Calibration Errors']).to_excel(writer, "Calibration Outputs", startcol = 6, startrow = 2, index = False, header = False)
    pd.DataFrame(['Hagan Prices']).to_excel(writer, "Calibration Outputs", startcol = 6, startrow = 19, index = False, header = False)
    pd.DataFrame(['Market Prices']).to_excel(writer, "Calibration Outputs", startcol = 6, startrow = 36, index = False, header = False)
    
    pd.DataFrame(['Calibration Results']).to_excel(writer, "Calibration Outputs", 
                                                   startcol = 0, 
                                                   startrow = 1, index = False, header = False)
    
    pd.DataFrame(['Parameter', 'fZero', 'gamma', 
                  'a', 'b', 'c', 'd', 'eta', 'delta']).to_excel(writer, 
                                                            "Calibration Outputs", 
                                                            startcol = 0, 
                                                            startrow = 2, index = False, header = False)
    
    pd.DataFrame(resultsCalibration).to_excel(writer,
                                            "Calibration Outputs", 
                                             startcol = 1, 
                                             startrow = 3, index = False, header = False)
 
    pd.DataFrame(errors, 
             index = np.unique(expiries), 
             columns = np.unique(tenors)).to_excel(writer, 
                              "Calibration Outputs", 
                              startcol = 3, # Column F
                              startrow = 3, # Line 4
                              index = True, 
                              header = True)
    pd.DataFrame(Chi2Prices, 
                 index = np.unique(expiries), 
                 columns = np.unique(tenors)).to_excel(writer, 
                                  "Calibration Outputs", 
                                  startcol = 3, # Column F
                                  startrow = 20, # Line 21
                                  index = True, 
                                  header = True)
                                                       
    pd.DataFrame(normalRectPrices, 
                 index = np.unique(expiries), 
                 columns = np.unique(tenors)).to_excel(writer, 
                                  "Calibration Outputs", 
                                  startcol = 3, # Column F
                                  startrow = 37, # Line 38
                                  index = True, 
                                  header = True)
                                                       
    pd.DataFrame(['1Y', '5Y', '10Y', '20Y']).transpose().to_excel(writer, 
                                        'Rates Distribution',
                                              startcol = 0, # Column F
                                            startrow = 0, # Line 38
                                            index = False, 
                                            header = False)

    pd.DataFrame(finalRatesDistributions[0], 
                 index = quantileIndex).to_excel(writer, 
                                            'Rates Distribution',
                                              startcol = 1, # Column F
                                                startrow = 2, # Line 38
                                                index = True, 
                                                header = True)

    pd.DataFrame(finalRatesDistributions[1], 
                 index = quantileIndex).to_excel(writer, 'Rates Distribution',
                                              startcol = 1, # Column F
                                                startrow = 19, # Line 38
                                                index = True, 
                                                header = True) 

    pd.DataFrame(finalRatesDistributions[2], 
                 index = quantileIndex).to_excel(writer, 'Rates Distribution',
                                              startcol = 1, # Column F
                                                startrow = 36, # Line 38
                                                index = True, 
                                                header = True)

    pd.DataFrame(finalRatesDistributions[3], 
                 index = quantileIndex).to_excel(writer, 'Rates Distribution',
                                              startcol = 1, # Column F
                                                startrow = 53, # Line 38
                                                index = True, 
                                                header = True) 
   
    'Save the variation in rates'
    '--------------------------------------'
    pd.DataFrame(['1Y', '5Y', '10Y', '20Y']).transpose().to_excel(writer, 
                                            'Rates Variation',
                                                  startcol = 0, # Column F
                                                startrow = 0, # Line 38
                                                index = False, 
                                                header = False)
    
    pd.DataFrame(ZCvariations[0], 
                 index = quantileIndex).to_excel(writer, 
                                            'Rates Variation',
                                              startcol = 1, # Column F
                                                startrow = 2, # Line 38
                                                index = True, 
                                                header = True)

    pd.DataFrame(ZCvariations[1], 
                 index = quantileIndex).to_excel(writer, 'Rates Variation',
                                              startcol = 1, # Column F
                                                startrow = 19, # Line 38
                                                index = True, 
                                                header = True) 

    pd.DataFrame(ZCvariations[2], 
                 index = quantileIndex).to_excel(writer, 'Rates Variation',
                                              startcol = 1, # Column F
                                                startrow = 36, # Line 38
                                                index = True, 
                                                header = True)

    pd.DataFrame(ZCvariations[3], 
                 index = quantileIndex).to_excel(writer, 'Rates Variation',
                                              startcol = 1, # Column F
                                                startrow = 53, # Line 38
                                                index = True, 
                                                header = True)     
                                                 
    'Save the proportion of rates'
    '--------------------------------------'
    pd.DataFrame(['1Y', '5Y', '10Y', '20Y']).transpose().to_excel(writer, 
                                        'Rates Proportion',
                                              startcol = 0, # Column F
                                            startrow = 0, # Line 38
                                            index = False, 
                                            header = False)
    
    pd.DataFrame(negativeRatesProportion[0], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 
                                            'Rates Proportion',
                                              startcol = 1, # Column F
                                                startrow = 2, # Line 38
                                                index = True, 
                                                header = True)

    pd.DataFrame(negativeRatesProportion[1], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 'Rates Proportion',
                                              startcol = 1, # Column F
                                                startrow = 16, # Line 38
                                                index = True, 
                                                header = True) 

    pd.DataFrame(negativeRatesProportion[2], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 'Rates Proportion',
                                              startcol = 1, # Column F
                                                startrow = 30, # Line 38
                                                index = True, 
                                                header = True)

    pd.DataFrame(negativeRatesProportion[3], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 'Rates Proportion',
                                              startcol = 1, # Column F
                                                startrow = 44, # Line 38
                                                index = True, 
                                                header = True)  
    
    writer.save()
    
    

'2) BULK RUN'
'====================================='    
def Run():
    # Obtain the folders in the Inputs Directory
    inputFolders = [x[0] for x in os.walk('Bulk Run')][1:]
    bulkRun = wb.sheets['Bulk Run']
    
    inputs = bulkRun.range('I8').expand('down').value 
    
    maxMaturity = 40
    maxProjYear = 50
    
    for directory in inputFolders:
        # Read Inputs
        curve = pd.read_excel(directory + '\\EIOPA Curve.xlsx',
                              header = None,
                              index_col=0).values.flatten()
        
        esgaviva.volatilitiesLMMPlus = pd.read_excel(directory + '\\Volatility Surface.xlsx', 
                                index_col = 0).unstack().reset_index(name='value')
        esgaviva.weightsLMMPlus = pd.read_excel(directory + '\\Weights.xlsx',
                                index_col = 0).unstack().reset_index(name='value')
        
        'Rate input treatment'
        '----------------------------'
        # Convert Rates to ZC Prices {(1/(1 + rate))^maturity}
        esgaviva.zeroCouponCurve = np.array([pow(1/(1+curve[i-1]), i) for i in range(1, len(curve)+1)])
        
        # Obtain the vol surface {(ZC(Tk-1)/ZC(Tk))-1}
        esgaviva.forwardCurve = np.array([(esgaviva.zeroCouponCurve[i-1]/esgaviva.zeroCouponCurve[i])-1 
                                              for i in range(1, len(esgaviva.zeroCouponCurve))])
        
        
        'Volatility and rate treatment'
        '----------------------------------'
        esgaviva.volatilitiesLMMPlus.columns = ['Tenor', 'Expiry', 'Value']
        esgaviva.weightsLMMPlus.columns = ['Tenor', 'Expiry', 'Value']
        
        # Unintuitive naming but used for Martingale and Monte Carlo Tests
        esgaviva.expiriesBachelier = esgaviva.weightsLMMPlus['Expiry'].to_numpy() 
        esgaviva.tenorsBachelier = esgaviva.weightsLMMPlus['Tenor'].to_numpy()

        # Calculate the forward Swap Rates for each Expiry X Tenor couple
        esgaviva.strikesLMMPlus = esgaviva.forwardSwapRateVect(esgaviva.weightsLMMPlus['Expiry'],
                                                               esgaviva.weightsLMMPlus['Tenor'] )

        # Calculate the Bachelier Price
        esgaviva.normalPricesLMMPlus = esgaviva.normalPayerVect(1, 
                                                                esgaviva.strikesLMMPlus, 
                                                                esgaviva.strikesLMMPlus,
                                                                esgaviva.weightsLMMPlus['Expiry'],
                                                                esgaviva.weightsLMMPlus['Tenor'],
                                                                esgaviva.volatilitiesLMMPlus['Value'])      
        # Perform all required treatments
        singleRun(inputs, 
                       directory + '\\DD LFM CEV YE19 Scenarios.csv', 
                       directory + '\\Audit Trail.xlsx', 
                       maxMaturity, maxProjYear)
        
        # Reset Variables
        'Not exactly the most optimal choice but many of the for loops are based on having default curves'
        '      in the module. Therefore, it was a fail stop at the last minute to help manage details as'
        '      opposed to cleaning up the entire set of code.'
        imp.reload(esgaviva)
    
    sg.popup('Run completed. Please refer to the Bulk Run Folder for Economic Scenario Tables.')
    
    
    
    
    
'3)  MULTISTART CALIBRATION'
'==============================================='   
def MultiStart():
    calibMultiStart = wb.sheets['Calibration - MultiStart']
    
    # Obtain all input parameters
    parameterTable = calibMultiStart.range('G10').expand('table').value   
    length = len(parameterTable)
    sg.popup(str(length) +' scenarios selected. Expected runtime: '+ str(length * 3) + ' minutes')
    
    results = []
    for paramSet in parameterTable: 
        results.append(esgaviva.haganCalibratorPython(paramSet[:6], 
                                                      eta = paramSet[6], 
                                                      delta = paramSet[7]))
        
    # Write results to Excel
    calibMultiStart.range('R10').value = results
    sg.popup('Multistart Calibration completed')
    
     
'4) CLOTURE'
'================================================'
def cloture():
    IM3Inputs = wb.sheets['IM3 Distributions']
    IM3Scenarios = wb.sheets['IM3 Scenarios']
    inputs = IM3Scenarios.range('I8').expand('down').value 
    
    # Principal components and distributions will be used in both cases
    IRPrincipalComponents = np.transpose(IM3Inputs.range('G29').expand('table').value) # Rates PComponents
    
    # Distributions
    distribValues = [0, 0.005, 0.01, 0.05, 0.1, 0.25, 0.5, 0.75, 0.9, 0.95, 0.99, 0.995, 1]
    riskFactors = ['FL', 'FS', 'FT', 'FVol', 'VACorp', 'VASov']
    
    distributions =  pd.DataFrame(IM3Inputs.range('G37').expand('table').value, 
                                    columns = riskFactors, 
                                    index = distribValues)
        
    if (IM3Scenarios.range('I17').value == 'True'):
        'INPUT MANAGEMENT'
        '------------------------'
        # Same names as folders in Cloture        
        scenarioNames = IM3Scenarios.range('G18').expand('down').value
        
        # Quantiles for IR Scenarios (Central, FL, FS, FT, FLUp, FSUp, FTUp)
        rateScenarios = [[0,0,0],
                         [distributions['FL'].loc[0.995], 0, 0],
                         [0, distributions['FS'].loc[0.995], 0],
                         [0, 0, distributions['FT'].loc[0.5]],
                         [distributions['FL'].loc[0.005], 0, 0],
                         [0, distributions['FS'].loc[0.005], 0],
                         [0, 0, distributions['FT'].loc[0.005]]]
        
        # Modify VA Scenarios
        VAScenarios = [distributions['VACorp'].loc[0.995], distributions['VASov'].loc[0.995]]
        
        
        'INTEREST RATE SCENARIO MODIFICATION'
        '=========================================='    
        # Do Runs for IR PCA Scenarios
        for i in range(len(scenarioNames[:7])):
            # Stress the Rate Curve
            defaultCurve = np.power(1/esgaviva.zeroCouponCurve[:40], 1/np.arange(1, 41))-1
            defaultCurve += np.dot(IRPrincipalComponents, rateScenarios[i]) # Use PCA shocks
            defaultZC = np.power(1/(1+defaultCurve), np.arange(1, 41)) # Convert to ZC
            
            # Apply Smith Wilson Extrapolation
            zcCurve = esgaviva.SmithWilson(defaultZC, alpha = 0.13681,  UFR = 0.0375).curve()
            forwardCurve = (zcCurve[:len(zcCurve)-1]/zcCurve[1:])-1
            
            esgaviva.zeroCouponCurve = zcCurve
            esgaviva.forwardCurve = forwardCurve
            
            # Run scenarios
            singleRun(inputs, 
                      scenariosAddress = 'Cloture\\'+scenarioNames[i]+'\\Scenario Table.csv', 
                      auditTrailAddress = 'Cloture\\'+scenarioNames[i]+'\\Audit Trail.xlsx', 
                      maxMaturity = 40, 
                      maxProjYear = 50)
            
            # Reset values (refer to line 516)
            imp.reload(esgaviva) #Have to do this to reset all vol & curve inputs
              
        
        'VA SCENARIO MODIFICATION'
        '=========================================='
        # Do Runs for IR Scenarios
        for j in range(len(scenarioNames[8:])):
            curve = np.power(1/esgaviva.zeroCouponCurve, 1/np.arange(1, 151))-1
            
            '*****Take note of the 0.7 haircut*******'
            curve += curve + 0.7 * VAScenarios[j]/10000 
            esgaviva.zeroCouponCurve = np.power(1/(1+curve), np.arange(1, 151))
            esgaviva.forwardCurve = (esgaviva.zeroCouponCurve[:149]/esgaviva.zeroCouponCurve[1:]) -1
            
            # Run scenarios
            singleRun(inputs, 
                      scenariosAddress = 'Cloture\\'+scenarioNames[j+8]+'\\Scenario Table.csv', 
                      auditTrailAddress = 'Cloture\\'+scenarioNames[j+8]+'\\Audit Trail.xlsx', 
                      maxMaturity = 40, 
                      maxProjYear = 50)
            
            # Reset values (refer to line 516)
            imp.reload(esgaviva) #Have to do this to reset all vol & curve inputs
        
            
        'VOL SURFACE MODIFICATION'
        '=========================================='
        # Volatility Surface
        FVolSurface = np.array(IM3Inputs.range('G11').expand('table').value) * distributions['FVol'].loc[0.995]/10000
        tenors = np.array(IM3Inputs.range('G9').expand('right').value).astype(int)
        expiries = np.array(IM3Inputs.range('E11').expand('down').value).astype(int)
        FVolSurface = pd.DataFrame(FVolSurface, 
                                   columns = tenors, 
                                   index = expiries)
        
        # Set volatility surface to the same structure as weights
        FVolSurface = FVolSurface.loc[np.unique(esgaviva.expiriesBachelier)].unstack().reset_index(name='value')
        FVolSurface.columns = ['Tenor', 'Expiry', 'Value']
        esgaviva.volatilitiesLMMPlus['Value'] += FVolSurface['Value']
        
        # Run scenarios
        singleRun(inputs, 
                  scenariosAddress = 'Cloture\\FVol\\Scenario Table.csv', 
                  auditTrailAddress = 'Cloture\\FVol\\Audit Trail.xlsx', 
                  maxMaturity = 40, 
                  maxProjYear = 50)
        
        # Reset values (refer to line 516)
        imp.reload(esgaviva) #Have to do this to reset all vol & curve inputs
        sg.popup('Run completed. Please check the:', '\'Cloture\' folder')
        
    else:
       PCdist =  IM3Scenarios.range('I30').expand('down').value
       PCScenarios = [distributions['FL'].loc[PCdist[0]],
                      distributions['FS'].loc[PCdist[1]],
                      distributions['FT'].loc[PCdist[2]]]
       
       # Stress the Rate Curve
       defaultCurve = np.power(1/esgaviva.zeroCouponCurve[:40], 1/np.arange(1, 41))-1
       defaultCurve += np.dot(IRPrincipalComponents, PCScenarios) # Use PCA shocks
       defaultZC = np.power(1/(1+defaultCurve), np.arange(1, 41)) # Convert to ZC
       
       # Apply Smith Wilson Extrapolation
       zcCurve = esgaviva.SmithWilson(defaultZC, alpha = 0.13681,  UFR = 0.375).curve()
       forwardCurve = (zcCurve[:len(zcCurve)-1]/zcCurve[1:])-1
       
       esgaviva.zeroCouponCurve = zcCurve
       esgaviva.forwardCurve = forwardCurve
        
       # Run scenarios
       singleRun(inputs, 
                  scenariosAddress = 'Cloture\\User Defined Distributions\\Scenario Table.csv', 
                  auditTrailAddress = 'Cloture\\User Defined Distributions\\Audit Trail.xlsx', 
                  maxMaturity = 40, 
                  maxProjYear = 50)
        
       # Reset values (refer to line 516)
       imp.reload(esgaviva) #Have to do this to reset all vol & curve inputs
       sg.popup('Run completed. Please check the:', '\'Cloture\\User Defined Distributions\\ \' folder')
      