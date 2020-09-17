'NOTE'
'---------------'
'* All computation is done within the esgaviva module'
'* This allows for modifying the module variables and functions'

# Manage operating systems
import os
import numpy as np
import pandas as pd
import PySimpleGUI as sg
from time import time
import imp

# Set Working Directory
loc = '//cerata/DGCU/Risques/16. Actuarial Function/05. Economic Scenario Générator/00 - DD LMM CEV/02 - Travaux/Excel Interface'
os.chdir(loc)

# Import aviva esg functions
import esgaviva 

#wb = esgaviva.xlw.Book.caller()
wb = esgaviva.xlw.Book('IR Nominal ESG Tool.xlsm')

sg.theme('DarkGrey2')

'1) Changing the model curve'
'------------------------------'
def changeRateCurve():
    'This function changes the module curve depending on the choice between inputing a manual curve and NSS parameters.'
    
    '**** This function will not be run using any button but at the beginning of each calibration/projection, will be used. ****'
    
    'Debug Assistance'
    '------------------'
    'i) Choice between NSS and Manual Curve:'
    'Sheet: Rate Curve - NSS - Choice'
    'Range: G8'
     
    'ii) Inputs (NSS):'
    'Sheet: Rate Curve - NSS - Own Params'
    'Range: I9 (going down)'
    
    # Define all the sheets to be used (Confirm these manually in the esgaviva2 excel fil)
    rateCurveNSSChoice = wb.sheets['Rate Curve - NSS - Choice']
    rateCurveInputs = wb.sheets['Rate Curve - Inputs']
    rateCurveOwnParams = wb.sheets['Rate Curve - NSS - Own Params']
    rateCurveResults = wb.sheets['Rate Curve - NSS - Results']
    rateCurveNSSOptimization = wb.sheets['Rate Curve - NSS - Optimization']
    
    if (rateCurveNSSChoice.range('G8').value == 'YES'):
        if (rateCurveNSSChoice.range('G9').value == 'YES'):
            #Obtain all the inputs for the NSS (I9 of the  Rate Curve NSS Choice sheet)
            inputs = rateCurveOwnParams.range('I9').expand().value 
            
            # Calibrate an NSS
            curveFunction = esgaviva.NelsonSiegelSvenssonCurve(beta0 = inputs[2],
                                                     beta1 = inputs[3],
                                                     beta2 = inputs[4],
                                                     beta3 = inputs[5],
                                                     tau1 = inputs[0],
                                                     tau2 = inputs[1])
            
            # Run NSS with input parameters
            newCurve = curveFunction(np.linspace(1, 150, 150))
            print('Please note that the curve selected is the NSS Curve') # Will probably change this to a messagebox
            
            # Write the results of the NSS Calibration to the Comparison Sheet
            rateCurveResults.range('I9').expand('down').value = [[elem] for elem in newCurve]
            sg.popup('NSS with own parameters selected')
            
        else:
            maturities = rateCurveInputs.range('E9:E158').value
            rates = rateCurveInputs.range('G9:G158').value
            
            # Find where rates are missing
            maturities = np.array(maturities)[np.where(np.array(rates) != None)[0]]
            rates = np.array(list(filter(None, rates)))
            
            # Calibrate NSS
            curve, status = esgaviva.calibrate_nss_ols(maturities, rates)
            
            # Obtain new Curve
            newCurve = curve(np.linspace(1, 150, 150))
            
            # Write parameters to Excel File
            params = [curve.beta0, curve.beta1, curve.beta2, curve.beta3, curve.tau1, curve.tau2, status.fun, status.success]
        
            # Write the results of the NSS Calibration to the Comparison Sheet
            rateCurveResults.range('I9').expand('down').value = [[elem] for elem in newCurve]
            rateCurveNSSOptimization.range('I13').value = [[param] for param in params]
      
    else:
        # Obtain the interest rate curve manually input
        newCurve = rateCurveInputs.range('G9').expand().value
        rateCurveResults.range('I9').expand('down').value = [[elem] for elem in newCurve]


    # If the inputCurve is too short, extrapolate using NSS        
    if (len(newCurve) < 90):
        maturities = rateCurveInputs.range('E9:E158').value
        rates = rateCurveInputs.range('G9:G158').value
        
        # Find where rates are missing
        maturities = np.array(maturities)[np.where(np.array(rates) != None)[0]]
        rates = np.array(list(filter(None, rates)))
        
        # Calibrate NSS
        curve, status = esgaviva.calibrate_nss_ols(maturities, rates)
        
        # Obtain new Curve
        newCurve = curve(np.linspace(1, 150, 150))
        
        # Write parameters to Excel File
        params = [curve.beta0, curve.beta1, curve.beta2, curve.beta3, curve.tau1, curve.tau2, status.fun, status.success]
    
        # Write the results of the NSS Calibration to the Comparison Sheet
        rateCurveResults.range('I9').expand('down').value = [[elem] for elem in newCurve]
        rateCurveNSSOptimization.range('I13').value = [[param] for param in params]
        

        
    # Transform curve to forwardCurve & zeroCouponCurve
    global  zeroCouponCurve, forwardCurve
    
    'Change module forwardCurve and zeroCouponCurve'
    zeroCouponCurve = np.power(1/(1+np.array(newCurve)), np.arange(1, 151))
    forwardCurve = (zeroCouponCurve[:len(zeroCouponCurve)-1]/zeroCouponCurve[1:])-1
    
    # Change the values in the esgaviva model
    esgaviva.zeroCouponCurve = zeroCouponCurve
    esgaviva.forwardCurve = forwardCurve 
    
    # Save Curve
    book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
    writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    
    pd.DataFrame(zeroCouponCurve).to_excel(writer, 
                                          "Initial Parameters", 
                                          startcol = 2, # Column F
                                          startrow = 1, # Line 59
                                          index = False, 
                                          header = False)
    
    pd.DataFrame(zeroCouponCurve).to_excel(writer, 
                                      "Martingale Test (CEV)", 
                                      startcol = 2, # Column F
                                      startrow = 3, # Line 59
                                      index = False, 
                                      header = False)
    
    writer.save()


'2a) Fill volatility surface if YE19 selected'
'---------------------------------------------'
def fillDefaultData():
    # Define all sheets to be used
    defaultData = wb.sheets['Default Data']    

    vols = esgaviva.volatilitiesLMMPlus.pivot(index = 'Expiry', 
                                              columns = 'Tenor', 
                                              values = 'Value').to_numpy()
    
    prices = esgaviva.normalPricesLMMPlus.reshape((len(vols), len(vols[0])))
    
    weights = esgaviva.weightsLMMPlus.pivot(index = 'Expiry',
                                            columns = 'Tenor', 
                                            values = 'Value').to_numpy()
    defaultData.range('C3').value = vols
    defaultData.range('C18').value = prices
    defaultData.range('C33').value = weights


'2b) Change volatility surface'
'-----------------------------------------'
def changeVolSurface():
    '''This function changes the input volatility and associated weights. Note that this curve should be a normal volatility
        surface'''
    
    # Define all sheets to be used
    calibrationInputs = wb.sheets['Calibration - Inputs']
    calibrationMktVols = wb.sheets['Calibration - MktVols']
    calibrationWeights = wb.sheets['Calibration - Weights']
    
    
    if (calibrationInputs.range('I8') == 'MANUAL_INPUT'):
        # Obtain the market volatilities
        tenorsSurface = calibrationMktVols.range('G8').options(expand = 'right').value # Obtain tenors 
        expiriesSurface = calibrationMktVols.range('E10').options(expand = 'down').value # Obtain maturities
        volatilities = calibrationMktVols.range('G10').options(expand = 'table').value
        
        # Convert tenorsSurface and expiriesSurface to integers
        tenorsSurface = [int(i) for i in tenorsSurface]
        expiriesSurface = [int(i) for i in expiriesSurface]
        
        # Find the shape of the table
        R = len(volatilities)
        C = len(volatilities[0])
        
        # Obtain the weights
        weights = calibrationWeights.range((10, 7), (10 + R-1, 7 + C-1)).options(empty = 0).value
        
        
        'These should be set to the volatilities in the esgaviva'
        # Convert to dataframe and save as volatility and weights LMM 
        volatilitiesLMMPlus = pd.DataFrame(volatilities, 
                                    index = expiriesSurface, 
                                    columns = tenorsSurface).unstack().reset_index(name ='Volatility')
        
        weightsLMMPlus = pd.DataFrame(weights, 
                                    index = expiriesSurface, 
                                    columns = tenorsSurface).unstack().reset_index(name ='Volatility')
        
        # Name the columns
        volatilitiesLMMPlus.columns = ['Tenor', 'Expiry', 'Value']
        weightsLMMPlus.columns = ['Tenor', 'Expiry', 'Value']
    
        
        'PRICE CALCULATION'
        # Calculate the swaption prices using Chi2 
        expiriesBachelier = weightsLMMPlus['Expiry'].to_numpy(int) 
        tenorsBachelier = weightsLMMPlus['Tenor'].to_numpy(int)
    
        strikesLMMPlus = esgaviva.forwardSwapRateVect(expiriesBachelier, tenorsBachelier)
        
        # Calculate bachelier prices
        normalPricesLMMPlus = esgaviva.normalPayerVect(1, strikesLMMPlus, strikesLMMPlus,
                                              expiriesBachelier, tenorsBachelier, volatilitiesLMMPlus['Value'])
        
        'Assign values to esgaviva module'
        '----------------------------'
        esgaviva.volatilitiesLMMPlus = volatilitiesLMMPlus
        esgaviva.weightsLMMPlus = weightsLMMPlus
        esgaviva.expiriesBachelier = expiriesBachelier
        esgaviva.tenorsBachelier = tenorsBachelier
        esgaviva.strikesLMMPlus = strikesLMMPlus
        esgaviva.normalPricesLMMPlus = normalPricesLMMPlus 
    

    # Otherwise fill data from the default dataset     
    else:
         fillDefaultData()
        
    # Save Surface & Weights
    book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
    writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    
    esgaviva.volatilitiesLMMPlus.pivot(index = 'Expiry', columns = 'Tenor', values = 'Value').to_excel(writer, 
                                          "Initial Parameters", 
                                          startcol = 5, # Column F
                                          startrow = 1, # Line 59
                                          index = True, 
                                          header = True)
    
    esgaviva.weightsLMMPlus.pivot(index = 'Expiry', columns = 'Tenor', values = 'Value').to_excel(writer, 
                                      "Initial Parameters", 
                                      startcol = 5, # Column F
                                      startrow = 16, # Line 59
                                      index = True, 
                                      header = True)
    writer.save()


'3. Define the iteration functions'
'----------------------------------'

def iter_cb_Chi(params, iter, resid):
     
    calibrationResults = wb.sheets['Calibration - Results - Retriev']
    
    parameters = [params['fZero'].value, params['gamma'].value, 
                  params['a'].value, params['b'].value, 
                  params['c'].value, params['d'].value, np.sum(np.square(resid))]
    
    calibrationResults.range((15 +iter, 7)).value = parameters
     
 
     
def iter_cb_Hagan(params, iter, resid):
    
    calibrationResults = wb.sheets['Calibration - Results - Retriev']
    
    parameters = [params['fZero'].value, params['gamma'].value, 
                  params['a'].value, params['b'].value, 
                  params['c'].value, params['d'].value, np.sum(np.square(resid))]
    
    calibrationResults.range((15 +iter, 18)).value = parameters
    
    

'4) Inputs to the calibration parameters'
'-----------------------------------------'
# i)  Multistart Calibrator   
def miniMultiStart():
    calibMultiStart = wb.sheets['Calibration - MultiStart']
    calibrationInputs = wb.sheets['Calibration - Inputs']
    
    # Number of iterations
    length = int(np.minimum(np.maximum(calibrationInputs.range('I36').value, 2), 500))
    
    # Obtain all input parameters
    metaParams = calibrationInputs.range('I17:I18').value
    parameterTable = np.random.rand(length, 6)
    parameterTable = [list(param)+metaParams for param in parameterTable]
    calibMultiStart.range('G10').expand('table').value = 0
    calibMultiStart.range('G10').expand('table').value = parameterTable
    
    sg.popup(str(length) +' scenarios selected. Expected runtime: '+ str(length) + ' minutes')
    
    results = []
    for paramSet in parameterTable: 
        results.append(esgaviva.haganCalibratorPython(paramSet[:6], 
                                                      eta = paramSet[6], 
                                                      delta = paramSet[7],
                                                       iter_cb = None))
        
    # Write results to Excel
    calibMultiStart.range('R10').value = results
    sg.popup('Multistart Calibration completed. Results in the Multistart page.',"",
             'Please select optimal parameters for the Projection.')
    
    
# ii)  Main Calibrator      
def calibrator():
    # Select new interest rate curve
    changeRateCurve()
    
    # Choose the volatility surface 
    changeVolSurface()
    
    calibrationInputs = wb.sheets['Calibration - Inputs']
    rateCurveOwnParams = wb.sheets['Rate Curve - NSS - Own Params']
    calibrationResults = wb.sheets['Calibration - Results - Retriev']
    calibrationBacktestPrice = wb.sheets['Calibration - Backtest - Price']
    calibrationBacktestVols = wb.sheets['Calibration - Backtest - Vols']
    
    # To be used in the projection section
    global eta, delta
    
    # Select inputs ie initialization, upper and lower bounds
    inputs = calibrationInputs.range('I11').expand().value
    upperBounds = calibrationInputs.range('I28').expand().value
    lowerBounds = calibrationInputs.range('I21').expand().value
    
    eta, delta = inputs[6], inputs[7]
    
    # Obtain the choic of calibration
    calibrationChoice = inputs[8]
    
    # Clear the two result tables
    calibrationResults.range('G15').expand('table').clear()
    calibrationResults.range('R15').expand('table').clear()
    
    # Create the global variable containing the results
    global resultsCalibration
    
    # Run Multistart if selected
    if (calibrationInputs.range('I35').value ==  'YES'):
        miniMultiStart()
        
    else:
        # Run Calibration based on the options selected in Excel
        if (calibrationChoice == 'Chi2'):
            start = time()
            resultsCalibration = esgaviva.chiSquareCalibratorPython(initialValues = inputs[:6], 
                                                 lowerBounds = lowerBounds, 
                                                 upperBounds = upperBounds, 
                                                 eta = inputs[6],
                                                 delta = inputs[7], 
                                                 iter_cb = iter_cb_Chi)
            end = time()
            calibrationResults.range('G11').value = resultsCalibration[:6]
            
            calibrationResults.range('M10').value = (end - start)/60
            
            # Save Surface & Weights
            book = esgaviva.load_workbook('Results\Results Summary.xlsx')
            writer = pd.ExcelWriter('Results\Results Summary.xlsx', engine = 'openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            
            pd.DataFrame(inputs[:8]).to_excel(writer, 
                                                  "Calibration Outputs", 
                                                  startcol = 1, # Column F
                                                  startrow = 2, # Line 59
                                                  index = False, 
                                                  header = False)
            
            pd.DataFrame(resultsCalibration[:6].extend([eta, delta])).to_excel(writer, 
                                                  "Calibration Outputs", 
                                                  startcol = 2, # Column F
                                                  startrow = 2, # Line 59
                                                  index = False, 
                                                  header = False)
            writer.save()
            
        elif (calibrationChoice == 'Hagan '):
            start = time()
            resultsCalibration = esgaviva.haganCalibratorPython(inputs[:6], 
                                                 upperBounds, 
                                                 lowerBounds, 
                                                 eta = inputs[6],
                                                 delta = inputs[7],
                                                 iter_cb = iter_cb_Hagan)
            end = time()
            
            calibrationResults.range('R11').value = resultsCalibration[:6]
            calibrationResults.range('X10').value = (end - start)/60
            
            # Save Surface & Weights
            book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
            writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            
            pd.DataFrame(inputs[:8]).to_excel(writer, 
                                                  "Calibration Outputs", 
                                                  startcol = 1, # Column F
                                                  startrow = 2, # Line 59
                                                  index = False, 
                                                  header = False)
            
            pd.DataFrame(resultsCalibration).to_excel(writer, 
                                                  "Calibration Outputs", 
                                                  startcol = 3, # Column F
                                                  startrow = 2, # Line 59
                                                  index = False, 
                                                  header = False)
            writer.save()
            
        else:
            start = time()
            resultsComparative = esgaviva.fullCalibratorPython(inputs[:6], 
                                upperBounds, 
                                 lowerBounds, 
                                 inputs[6],
                                 inputs[7], 
                                 iter_chi = iter_cb_Chi,
                                 iter_Hagan = iter_cb_Hagan)
            end = time()
            calibrationResults.range('X10').value = (end - start)/60
            
            # The default choice is the Chi2
            resultsCalibration = resultsComparative[1]
            calibrationResults.range('G11').value = resultsComparative[0][:6]
            calibrationResults.range('R11').value = resultsComparative[1][:6]
            
            # Save Surface & Weights
            book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
            writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            
            pd.DataFrame(inputs[:8]).to_excel(writer, 
                                                  "Calibration Outputs", 
                                                  startcol = 1, # Column F
                                                  startrow = 2, # Line 59
                                                  index = False, 
                                                  header = False)
            
            pd.DataFrame(resultsComparative[0][:6].extend([eta, delta])).to_excel(writer, 
                                                  "Calibration Outputs", 
                                                  startcol = 2, # Column F
                                                  startrow = 2, # Line 59
                                                  index = False, 
                                                  header = False)
            
            pd.DataFrame(resultsComparative[1][:6].extend([eta, delta])).to_excel(writer,
                                              "Calibration Outputs", 
                                              startcol = 3, # Column F
                                              startrow = 2, # Line 59
                                              index = True, 
                                              header = True)
            writer.save()
        
        'PRICE CALCULATION MODEL'
        '----------------------------'
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
        errors = errors.reshape((int(rows), cols)).T
        Chi2Prices = chiSquarePrices.reshape((int(rows), cols)).T
        normalRectPrices = normalPrices.reshape((int(rows), cols)).T
        
        calibrationBacktestPrice.range('G10').options(np.array, expand = 'table').value = errors
        calibrationBacktestPrice.range('G28').options(np.array, expand = 'table').value = Chi2Prices 
    
        # Save to Summary File
        book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
        writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)  
        
        pd.DataFrame(errors, 
                     index = np.unique(expiries), 
                     columns = np.unique(tenors)).to_excel(writer, 
                                      "Calibration Outputs", 
                                      startcol = 5, # Column F
                                      startrow = 3, # Line 4
                                      index = True, 
                                      header = True)
        pd.DataFrame(Chi2Prices, 
                     index = np.unique(expiries), 
                     columns = np.unique(tenors)).to_excel(writer, 
                                      "Calibration Outputs", 
                                      startcol = 5, # Column F
                                      startrow = 20, # Line 21
                                      index = True, 
                                      header = True)
                                                           
        pd.DataFrame(normalRectPrices, 
                     index = np.unique(expiries), 
                     columns = np.unique(tenors)).to_excel(writer, 
                                      "Calibration Outputs", 
                                      startcol = 5, # Column F
                                      startrow = 37, # Line 38
                                      index = True, 
                                      header = True)
                                                         
        
        writer.save()
        sg.popup('Calibration Completed')
        
        'BACHELIER VOLATILITY CALCULATION'
        '------------------------------------'
        # Turn Prices to volatilities
        bachelierModelVolatilities = esgaviva.volNormalATMFunctionVect(expiries, tenors, chiSquarePrices)
        volErrors =  np.abs((esgaviva.volatilitiesLMMPlus['Value']- bachelierModelVolatilities)/
                            esgaviva.volatilitiesLMMPlus['Value'])
        volErrors = volErrors.to_numpy().reshape((int(rows), cols)).T
        bachVols = bachelierModelVolatilities.reshape((int(rows), cols)).T
        
        calibrationBacktestVols.range('G10').options(np.array, expand = 'table').value = volErrors
        calibrationBacktestVols.range('G28').options(np.array, expand = 'table').value = bachVols
        

'5) Monte Carlo Pricing'
'-----------------------------'
def MonteCarlo():
    # Change Rate Curve
    changeRateCurve()
    
    # Choose the volatility surface 
    changeVolSurface()
    
    MCInputs = wb.sheets['MC - Inputs']
    MCResultsPrix = wb.sheets['MC - Results - Price']
    MCResultsVols = wb.sheets['MC - Results - Vols']
    calibrationInputs = wb.sheets['Calibration - Inputs']
    calibrationResults = wb.sheets['Calibration - Results - Retriev']
    
    # Obtain the vector of inputs
    inputs = MCInputs.range('I8').expand('down').value
    
    
    # Eta and delta declared
    resultsCalibration = calibrationResults.range('R11').expand('right').value
    eta, delta = calibrationInputs.range('I17').value, calibrationInputs.range('I18').value
    
    simus = int(inputs[1])
    
    # Run Monte Carlo Simulation depending on the case chosen
    if (inputs[5] == 'Python'): 
        MCPrices = [esgaviva.MonteCarloPricer(forwardCurve, expiry, tenor, resultsCalibration[0], 
                         resultsCalibration[1], 
                         resultsCalibration[2], 
                         resultsCalibration[3],
                         resultsCalibration[4], 
                         resultsCalibration[5], eta, esgaviva.betas, delta, simus) for 
                        expiry, tenor in list(zip(esgaviva.expiriesBachelier, 
                                                  esgaviva.tenorsBachelier))]
    
    elif (inputs[5] == 'YE19_BH'): 
        browniansBH = pd.read_excel(esgaviva.gaussiansLocation, 
                                    names = ['Trial', 'Timestep', 'Gaussian1', 'Gaussian2'], 
                                    index_col = [0,1])
        maxCol = np.amax(np.array(browniansBH.index.get_level_values(1)))
        simulations = int(len(browniansBH)/maxCol)
        browniensBH1 = browniansBH['Gaussian1'].to_numpy().reshape((simulations, maxCol))
        browniensBH2 = browniansBH['Gaussian2'].to_numpy().reshape((simulations, maxCol))
        
        MCPrices = [esgaviva.MonteCarloPricerBH(esgaviva.forwardCurve, expiry, tenor, resultsCalibration[0], 
                         resultsCalibration[1], 
                         resultsCalibration[2], 
                         resultsCalibration[3],
                         resultsCalibration[4], 
                         resultsCalibration[5], 
                         eta, esgaviva.betas, delta, browniensBH1, browniensBH2) for 
                        expiry, tenor in list(zip(esgaviva.expiriesBachelier, 
                                                  esgaviva.tenorsBachelier))]
    else:
        sg.popup('Under Construction')
        
    # Reshape MC Prices
    cols = len(esgaviva.weightsLMMPlus.pivot(index = 'Expiry', columns = 'Tenor', values = 'Value'))
    rows = len(MCPrices)/cols
    
    # Transform vols to prices
    MCVols = esgaviva.volNormalATMFunctionVect(esgaviva.weightsLMMPlus['Expiry'], 
                                      esgaviva.weightsLMMPlus['Tenor'], 
                                      MCPrices)
    
    # Find Errors
    MCPriceErrors = np.abs((esgaviva.normalPricesLMMPlus - MCPrices)/esgaviva.normalPricesLMMPlus)
    MCVolErrors = np.abs((esgaviva.volatilitiesLMMPlus['Value'] - MCVols)/
                                 esgaviva.volatilitiesLMMPlus['Value'])
    
    # Write vols and price errors
    MCResultsPrix.range('G10').value = MCPriceErrors.reshape((int(rows), cols)).T
    MCResultsVols.range('G10').value = MCVolErrors.to_numpy().reshape((int(rows), cols)).T
    
    # Write Prices
    MCResultsPrix.range('G28').value = np.array(MCPrices).reshape((int(rows), int(cols))).T
    MCResultsVols.range('G28').value = MCVols.reshape((int(rows), int(cols))).T
    
    # Save Surface & Weights
    book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
    writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    pd.DataFrame(np.array(MCPriceErrors).reshape((int(rows), int(cols))).T, 
                 index = np.unique(esgaviva.expiriesBachelier),
                 columns = np.unique(esgaviva.tenorsBachelier)).to_excel(writer, 
                                                                         "Monte Carlo Test", 
                                                                            startcol = 1, # Column F
                                                                               startrow = 2, # Line 38
                                                                               index = True, 
                                                                               header = True)
    
    pd.DataFrame(np.array(MCPrices).reshape((int(rows), int(cols))).T, 
             index = np.unique(esgaviva.expiriesBachelier), 
             columns = np.unique(esgaviva.tenorsBachelier)).to_excel(writer, 
                                                                     "Monte Carlo Test", 
                                                                            startcol = 1, # Column F
                                                                               startrow = 17, # Line 38
                                                                               index = True, 
                                                                               header = True)
    
    pd.DataFrame(np.array(MCVols).reshape((int(rows), int(cols))).T, 
                 index = np.unique(esgaviva.expiriesBachelier), 
                 columns = np.unique(esgaviva.tenorsBachelier)).to_excel(writer, 
                                                                         "Monte Carlo Test", 
                                                                            startcol = 1, # Column F
                                                                               startrow = 32, # Line 38
                                                                               index = True, 
                                                                               header = True)
    
    writer.save()
    

'6. Projector'    
'------------------'    
def Projector():
    ProjectionInputs = wb.sheets['Projection - Inputs']
    ProjectionResultsDistrib = wb.sheets['Projection - Results - Distrib']
    ProjectionResultsDiscount = wb.sheets['Projection - Distrib - Discount']
    ProjectionResultsProp = wb.sheets['Projection - Results - Proport']
    MGZCPrices = wb.sheets['MG - ZC Prices']
    discountFactors = wb.sheets['MG - Discount factors']
    calibrationResults = wb.sheets['Calibration - Results - Retriev']
    calibrationInputs = wb.sheets['Calibration - Inputs']
    ProjectionZCDiff = wb.sheets['Projection - Distrib - ZCDiff']
       
    # Change input parameters
    if (all(x != None for x in ProjectionInputs.range('I8:I15').value)):
        resultsCalibration = ProjectionInputs.range('I8').expand('down').value
        eta, delta = resultsCalibration[6], resultsCalibration[7]
        sg.popup('Please note that manual projection parameters will be used.')
    else:
        resultsCalibration = calibrationResults.range('R11').expand('right').value
        eta, delta = calibrationInputs.range('I17').value, calibrationInputs.range('I18').value
        sg.popup('Please note that calibrated parameters will be used.')

    # Change rate curve if NSS parameters are input
    if (all(x != None for x in ProjectionInputs.range('I17:I22').value)):
        NSSInputs = ProjectionInputs.range('I8').expand('down').value
        curveFunction = esgaviva.NelsonSiegelSvenssonCurve(NSSInputs[0], NSSInputs[1], 
                                             NSSInputs[2], NSSInputs[3], 
                                             NSSInputs[4], NSSInputs[5])
        
        # Change the curve
        global newCurve, zeroCouponCurve, forwardCurve
        newCurve = curveFunction(np.linspace(1, 150))

        'Change module forwardCurve and zeroCouponCurve'
        esgaviva.zeroCouponCurve = np.power(1/(1+np.array(newCurve)), np.arange(1, 151))
        esgaviva.forwardCurve = (esgaviva.zeroCouponCurve[:len(zeroCouponCurve)-1]/esgaviva.zeroCouponCurve[1:])-1
    
    else:
        # Select new interest rate curve
        changeRateCurve()
    
        # Choose the volatility surface 
        changeVolSurface()
        
    # Obtain the parameters we will use in our projection
    projInput1 = ProjectionInputs.range('I24').expand('down').value
    maxMaturity = int(projInput1[0])
    maxProjYear = int(projInput1[1])
    simu =projInput1[2]
    
    # Output Management
    projInput2 = ProjectionInputs.range('I28').expand('down').value
    sortieBrowniens = projInput2[0]
    inputBrowniens = projInput2[1]
    inputBetas = projInput2[2]   
    
    'To be used in prophet calculation'
    # Calculate the swaption prices using Chi2 
    expiriesProphet = np.tile(np.linspace(1, 80, 80), 10).astype(int)
    tenorsProphet = np.repeat(np.linspace(1, 10, 10), 80).astype(int)   
    strikesProphet = esgaviva.forwardSwapRateVect(expiriesProphet, tenorsProphet)
    
    '''PROJECTION MODULE
    ============================='''
    if (inputBrowniens == 'YE19_BH'):
        '''Moodys projection
        ========================='''
        if (inputBetas == 'YE19_BH'):
            results = esgaviva.BHfullSimulator(esgaviva.forwardCurve, 
                        resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5], eta, esgaviva.betas, delta,
                        maxProjectionYear = maxProjYear, maxMaturity = maxMaturity)
            
        'Monte Carlo pricing'
        '-----------------------------'
        browniansBH = pd.read_excel(esgaviva.gaussiansLocation, 
                                    names = ['Trial', 'Timestep', 'Gaussian1', 'Gaussian2'], 
                                    index_col = [0,1])
        maxCol = np.amax(np.array(browniansBH.index.get_level_values(1)))
        simulations = int(len(browniansBH)/maxCol)
        browniensBH1 = browniansBH['Gaussian1'].to_numpy().reshape((simulations, maxCol))
        browniensBH2 = browniansBH['Gaussian2'].to_numpy().reshape((simulations, maxCol))
        
        MCPrices = esgaviva.calibrationBlackVect(1.0, strikesProphet, 
                        strikesProphet,
                        expiriesProphet, 
                        tenorsProphet,
                        resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5],
                        eta, delta)
        
    elif (inputBrowniens == 'Python'):
        '''Python projection
        ================================'''
        if (inputBetas == 'YE19_BH'):
            if (sortieBrowniens == 'True'):
                results = esgaviva.fullSimulator(esgaviva.forwardCurve, resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5],eta, esgaviva.betas, delta,
                        simulations = 3000, 
                        maxProjectionYear = maxProjYear, 
                        maxMaturity = maxMaturity)
                
                # Save Gaussians to Excel
                pd.DataFrame(results[2]).to_csv('Results/Gaussians Python1.csv')
                pd.DataFrame(results[3]).to_csv('Results/Gaussians Python2.csv')
            
            else:
                results = esgaviva.fullSimulator(esgaviva.forwardCurve, resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5], eta, esgaviva.betas, delta,
                        simulations = 3000, maxProjectionYear = 50, maxMaturity = 40)
            
            'Monte Carlo pricing'
            '-----------------------------'
            MCPrices = esgaviva.calibrationBlackVect(1.0, strikesProphet, 
                        strikesProphet,
                        expiriesProphet, 
                        tenorsProphet,
                        resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5],
                        eta, delta)
    
        
    else:
        '''Sobol Projection
        ======================='''
        results = esgaviva.sobolSimulator(esgaviva.forwardCurve, resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5],eta, esgaviva.betas, delta,
                        maxProjectionYear = maxProjYear, 
                        maxMaturity = maxMaturity)
        
        MCPrices = esgaviva.calibrationBlackVect(1.0, strikesProphet, 
                        strikesProphet,
                        expiriesProphet, 
                        tenorsProphet,
                        resultsCalibration[0],
                        resultsCalibration[1],
                        resultsCalibration[2],
                        resultsCalibration[3],
                        resultsCalibration[4],
                        resultsCalibration[5],
                        eta, delta)
    
    sg.popup('Scenarios generated.', '', 'Data processing in progress', '', 'Estimated time: 15mins')
    
    'Save Monte Carlo Volatilities'
    '-----------------------------'    
    bachelierModelVolatilitiesProphet = esgaviva.volNormalATMFunctionVect(expiriesProphet, 
                                                                          tenorsProphet,
                                                                          MCPrices)*100
    
    # Save to Excel        
    pd.DataFrame(bachelierModelVolatilitiesProphet.reshape((10, 80)), 
                 columns = np.linspace(1, 80, 80),
                 index = np.linspace(1, 10, 10)).transpose().to_excel('Results\Vol Implicites(Prophet).xlsx')
    
    
    'MODIFYING THE CAPECO SCENARIO TABLE '
    '----------------------------------------------'
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
            
    #Select the boundaries required
    ZCdistributions = np.zeros((maxMaturity*len(ZCTimeToMaturity), maxProjYear))
    ratesdistributions = np.zeros((maxMaturity*len(ZCTimeToMaturity), maxProjYear))
    for i in range(len(ZCTimeToMaturity)):
        for j in range(maxMaturity):
            ZCdistributions[i*maxMaturity+j] = np.array(ZCTimeToMaturity[i][j][:maxProjYear])
            ratesdistributions[i*maxMaturity+j] = np.power(1/ZCdistributions[i*maxMaturity+j], 1/(j+1))-1
        
    'SCENARIO MANAGEMENT'
    'Scenarios generated by Moodys are modified by ESG results from DD LFM CEV'    
    if (ProjectionInputs.range('I34').value == 'ALM Format Scenarios'):
        # Import Moody's scenarios
        sg.popup('Please ensure there is a csv file named: LMMPlusBachelier Scenarios.csv in the Excel Interface directory.', '', 
                 'The format of this file should exactly match that of the DD LFM CEV Scenarios.csv in the Results Folder.')
        
        data = esgaviva.dk.read_csv('LMMPlusBachelier Scenarios.csv').compute()
        maturities = [1, 2, 3, 4, 5, 10, 15, 20, 25, 30, 35, 40]
 
        # From data, set index to the parameter column
        data2 = data.set_index('Parameter', append = True)

        # Find all the indices we will need for each scenario
        indexProphet = [list(np.array(maturities)+i*maxMaturity -1) for i in range(len(simulatedCurves))]
        indexProphet = list(esgaviva.itertools.chain.from_iterable(indexProphet))
        
        zcToAdd = list(esgaviva.zeroCouponCurve[np.array(maturities)-1])*len(simulatedCurves)
        ratesToAdd = list(np.power(1/esgaviva.zeroCouponCurve[np.array(maturities)-1], 
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
            newDF = np.cumprod(finalZCdistributions[1*i-1])
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
            data2.loc[i] = subdata
        
        # Save Final Data as csv
        data2.to_csv('Results/DD LFM CEV YE19 Scenarios.csv', chunksize=100)
        sg.popup('ALM Format scenarios generated. Output file will be DD LFM CEV Scenarios.csv in the Results folder')
        
    elif (ProjectionInputs.range('I34').value == 'YE19 Scenarios'):
        data2 = pd.read_pickle('YE19 Scenarios.pkl')
        maturities = [1, 2, 3, 4, 5, 10, 15, 20, 25, 30, 35, 40]
 
        # From data, set index to the parameter column
        data2 = data2.set_index('Parameter', append = True)

        # Find all the indices we will need for each scenario
        indexProphet = [list(np.array(maturities)+i*maxMaturity-1) for i in range(len(simulatedCurves))]
        indexProphet = list(esgaviva.itertools.chain.from_iterable(indexProphet))
        
        zcToAdd = list(esgaviva.zeroCouponCurve[np.array(maturities)-1])*len(simulatedCurves)
        ratesToAdd = list(np.power(1/esgaviva.zeroCouponCurve[np.array(maturities)-1], 
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
        data2.to_csv('Results/DD LFM CEV YE19 Scenarios.csv', chunksize = 100)
        sg.popup('YE19 scenarios generated. Output file will be DD LFM CEV Scenarios.csv in the Results folder')
        
    else:
        # Maturities to be inputfrom
        maturities = ProjectionInputs.range('I35').expand('down').value
        maturities = [elem for elem in maturities if elem < maxMaturity]  
        
        # Find all the indices we will need for each scenario
        indexProphet = [list(np.array(maturities).astype(int)+i*maxMaturity-1) for i in range(len(simulatedCurves))]
        indexProphet = list(esgaviva.itertools.chain.from_iterable(indexProphet))
        
        zcToAdd = list(esgaviva.zeroCouponCurve[np.array(maturities).astype(int)-1])*len(simulatedCurves)
        ratesToAdd = list(np.power(1/esgaviva.zeroCouponCurve[np.array(maturities).astype(int)-1], 
                                   1/np.array(maturities).astype(int))-1)*len(simulatedCurves)
        
        # Filter our data and add the current ZC and rate values
        finalZCdistributions = [np.append(zc,curve) for zc, curve in list(zip(zcToAdd, ZCdistributions[indexProphet]))]
        finalRatesdistributions = [np.append(rate,curve) for rate, curve in list(zip(ratesToAdd, ratesdistributions[indexProphet]))]
        dataIndex = maturities*len(simulatedCurves) 
        
        # Save Rates to csv
        pd.DataFrame(finalRatesdistributions, index = dataIndex).to_csv('Results/Raw Rate Scenarios.csv', chunksize = 100)
        pd.DataFrame(finalZCdistributions, index = dataIndex).to_csv('Results/Raw ZC Scenarios.csv', chunksize = 100)
        sg.popup('Raw scenarios generated. Output files will be Raw Rate/Raw ZC Scenarios.csv in the Results folder')
    
    'OBTAIN THE DISTRIBUTIONS'
    '==================================================================================================='
    # Select the maturities we want to analyze
    distributionMaturities = [1, 5, 10, 20] # Years to be considered
    quantiles = [0.5, 1, 5, 10, 25, 50, 75, 90, 95, 99, 99.5] # Quantiles that interest us
    ranges = [-1,-0.2, -0.10, -0.05, -0.025, 0, 0.025, 0.05, 0.1, 0.2, 0.3, 1] # These will be the ranges for the negative rates table
    
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
        rates = np.array([np.append([esgaviva.EIOPACurve[maturity - 1]], elem) for elem in np.power(1/distribs, 1/maturity) -1])
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
            negRates.append(np.histogram(rates.T[i], ranges)[0]/len(simulatedCurves))
        
        negativeRatesProportion.append(np.transpose(negRates))
        

    'Save the rates distributions to Excel'
    '---------------------------------'
    ProjectionResultsDistrib.range((11,7)).expand('table').value = finalRatesDistributions[0]
    ProjectionResultsDistrib.range((28,7)).expand('table').value = finalRatesDistributions[1]
    ProjectionResultsDistrib.range((45,7)).expand('table').value = finalRatesDistributions[2]
    ProjectionResultsDistrib.range((62,7)).expand('table').value = finalRatesDistributions[3]
    
    
    'Save the ZC Distributions to Excel'
    '------------------------------------'
    ProjectionResultsDiscount.range((11,7)).expand('table').value = finalZCDistributions[0]
    ProjectionResultsDiscount.range((25,7)).expand('table').value = finalZCDistributions[1]
    ProjectionResultsDiscount.range((39,7)).expand('table').value = finalZCDistributions[2]
    ProjectionResultsDiscount.range((53,7)).expand('table').value = finalZCDistributions[3]

    
    'Save the Rate Proportions to Excel'
    '------------------------------------'
    ProjectionResultsProp.range((12,7)).expand('table').value = negativeRatesProportion[0]
    ProjectionResultsProp.range((26,7)).expand('table').value = negativeRatesProportion[1]
    ProjectionResultsProp.range((40,7)).expand('table').value = negativeRatesProportion[2]
    ProjectionResultsProp.range((54,7)).expand('table').value = negativeRatesProportion[3]    

    'Save the ZC Distributions to Excel'
    '------------------------------------'
    ProjectionZCDiff.range((11,7)).expand('table').value = ZCvariations[0]
    ProjectionZCDiff.range((28,7)).expand('table').value = ZCvariations[1]
    ProjectionZCDiff.range((45,7)).expand('table').value = ZCvariations[2]
    ProjectionZCDiff.range((62,7)).expand('table').value = ZCvariations[3]  
    
    'MARTINGALE TEST'
    '==================================================================================================='
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

    # Average of ZC Simulations
    ##################################################################################
    # Copy the shape of the 1st triangle of the ZCoupons simulations
    zcMean = esgaviva.copy.deepcopy(zCoupons[0])
    
    # For calculation of the mean ZCoupons
    for i in range(len(resMean)):
        zcMean[i] = [0]*len(zcMean[i])
    
    # Calculate the mean ZCoupon len(simulatedCurves) = number of simulations
    for i in range(len(zCouponsTilde)):
        for j in range(len(zCouponsTilde[0])):
            zcMean[j] = np.nansum([zcMean[j], zCoupons[i][j]], axis = 0)
    
    zCouponAvg = [i/len(simulatedCurves) for i in zcMean]
    
    
    'CONVERT TRIANGLE TO RECTANGLE'
    '===================================='
    # Average deflated ZC (Take the First 30 maturities  for each year projected)
    rectangleZCTildeAvg = [zCouponTildeAvg[i][:30] for i in range(50)]
    
    # First Simulation
    ZCSimu1 = [zCoupons[0][i][:30] for i in range(len(zCoupons[0]))]
    forwardsSimu1 = [simulatedCurves[0][i][:30] for i in range(len(simulatedCurves[0]))]

    # Error Calculation
    errorsMGTest = [abs((rectangleZCTildeAvg[i]/esgaviva.zeroCouponCurve[1+i:31+i]) -1)
                                for i in range(len(rectangleZCTildeAvg))]
    
    # Save results
    for i in range(len(errorsMGTest)):
        MGZCPrices.range((10+i, 7)).expand('right').value = errorsMGTest[i]
    
    discountFactors.range('G10').value = AvgDeflateur[:50]
    discountFactors.range('G11').value = esgaviva.zeroCouponCurve[:50]  
    
    # Save Surface & Weights
    book = esgaviva.load_workbook('Results\Audit Trail.xlsx')
    writer = pd.ExcelWriter('Results\Audit Trail.xlsx', engine = 'openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    
    pd.DataFrame(AvgDeflateur[:50]).to_excel(writer, 'Martingale Test (CEV)',
                                              startcol = 3, # Column F
                                                startrow = 3, # Line 38
                                                index = False, 
                                                header = False)

    pd.DataFrame(errorsMGTest).to_excel(writer, 'Martingale Test (CEV)',
                                              startcol = 6, # Column F
                                                startrow = 3, # Line 38
                                                index = False, 
                                                header = False)   
    
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
                                              startcol = 2, # Column F
                                                startrow = 2, # Line 38
                                                index = False, 
                                                header = True)

    pd.DataFrame(negativeRatesProportion[1], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 'Rates Proportion',
                                              startcol = 2, # Column F
                                                startrow = 16, # Line 38
                                                index = True, 
                                                header = True) 

    pd.DataFrame(negativeRatesProportion[2], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 'Rates Proportion',
                                              startcol = 2, # Column F
                                                startrow = 30, # Line 38
                                                index = False, 
                                                header = True)

    pd.DataFrame(negativeRatesProportion[3], 
                 index = ranges[:len(ranges)-1]).to_excel(writer, 'Rates Proportion',
                                              startcol = 2, # Column F
                                                startrow = 44, # Line 38
                                                index = False, 
                                                header = True)  
    
    
    writer.save()
    imp.reload(esgaviva)
    sg.popup('Projection Completed')