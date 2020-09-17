                       AVIVA DD LFM CEV SCENARIO GENERATOR
=======================================================================================

            GENERAL INFORMATION
        ===========================
PLEASE BE SURE TO READ THIS SECTION BEFORE USE.

This is the interactive portal for the DD LFM CEV scenario generator.

The bulk of the code is written in the CalibSetup.py file. 

In case of any errors or troubleshooting please contact the France Actuarial Function.

For versioning changes or updates, please send a log of the updates to the France Actuarial Function.

            BASIC REQUIREMENTS
        =========================
1) Functioning Python version (version <3.6)

2) Functioning default R (should be installed as the default version)

3) Visit: https://confluence.aviva.co.uk/pages/viewpage.action?pageId=308706577 to set up Python if importing 
	packages fails


            INPUT & OUTPUT MANAGEMENT
        ===============================
1) Inputs
------------
- All general inputs for the model are contained in the Input Folder.
- The specific names are VERY important. Therefore, while changing an input, please be sure to change the name.

Annex Folder
----------------
- You can save any supplementary inputs in the Annex folder

2) Outputs
----------------
************ All Results are saved in the Outputs Folder ************


            MAIN FUNCTIONS
        ===========================

Function List
---------------

* fullCalibrator(initialValues, lowerBounds, upperBounds, eta, delta)

* fullSimulator(forwardCurve, fZero, gamma, a, b, c, d, eta, betas, delta, 
                  simulations = 5000, 
                  maxProjectionYear = 50, 
                  maxMaturity = 40,
                  viewDistributions = True,
                  zeroCouponMGTest = True, 
                  marketConsistencyTest = True, 
                  exportGaussians = True)

* BHfullSimulator(forwardCurve, fZero, gamma, a, b, c, d, eta, betas, delta,  
                  maxProjectionYear = 50, 
                  maxMaturity = 40,
                  viewDistributions = True,
                  zeroCouponMGTest = True, 
                  marketConsistencyTest = True)

        
1) Calibrator Functions
####################################

- These functions are the calibration workhorse. They perform the calibration given the input curve & surface.
- Two of these functions exist: chiSquareCalibratorR and chiSquareCalibratorPython.
- The complete function is the fullCalibrator function that combines the two

INPUTS
---------

**** ALL INPUTS SHOULD BE IN NUMERIC FORMAT i.e add .0 to every integer ****
a) Initial Values
- The input should be a 6 element list of the form [0,0,0,0,0,0]
- Each of these values represent [fZero, gamma, a, b, c, d]

b) Bounds 
- The boundaries for each of the inputs is also set in a list of the same length:
- [0.0,0.0,0.0,0.0,0.0,0.0] is the default Lower bound
- [1.0,1.0,1.0,1.0,1.0,1.0] is the default upper bound

c) Eta (elasticity) and delta (shift)
- Optional Values with set defaults at 0.8 and 0.1 respectively

Example:
 # Calibrating using the R minpack package
chiSquareCalibratorR(initialValues = [0.1,0.1,0.1,0.1,0.1,0.1], 
                     lowerBounds = [0.0,0.0,0.0,0.0,0.0,0.0], 
                     upperBounds = [1.0,1.0,1.0,1.0,1.0,1.0], 
                     eta = 0.8, 
                     delta = 0.1):

OUTPUTS
----------
- An optimized list of parameters. This can be plugged into the simulator functions.
- To compare the two, please check the Calibration sheet of the Results.xlsx file in the Outputs folder

    
2) fullSimulator Functions
####################################
    
- These are the main functions in the ESG. They allow for simulation of forward rate scenarios for a specific horizon.

- Two functions exist based on the choice of Gaussians i.e. Python gaussians (fullSimulator function) vs Input 
    Gaussians (BHfullSimulator function).
    
INPUTS
---------

a) Forward Curve
- The forward curve to be projected. Minimal length should be 80Y (Recommended length should be 90Y)
    
b) Calibration parameters
- Obtained from functions in (1) 

c) Projection Test Parameters
    i)   zeroCouponMGTest: Binary value to allow for zero Coupon Martingale Tests;
    
    ii)  marketConsistencyTest/SwaptionMonteCarloTest: Conduct all swaption pricing & martingale Tests;
            ****** Takes 3 - 5 minutes ******
            
    iii) exportGaussians: Export Gaussians used (Only for Python Calibrator);
    
    iv)  maxProjectionYear: Maximum number of years to project out to;
    
    v)   maxMaturity: Maximum maturity of Forwards projected out to maxProjectionYear eg 
            40Y curve (maxMaturity) projected out to 50 years (maxProjectionYear)
            
            *********** This will require a curve of length maxMaturity + maxProjectionYear *********

# Example
            
# Simulate using custom Gaussians
results =  BHfullSimulator(forwardCurve, fZero, gamma, a, b, c, d, eta, betas, delta, 
                    zeroCouponMGTest = True, 
                    SwaptionMonteCarloTest = True, 
                    viewDistributions = True,
                    maxProjectionYear = 50, 
                    maxMaturity = 40)           
            
            
            
            
            
OUTPUTS 
------------
i) All ZC scenarios are saved in the ZCScenarios.csv file. This should be directly utilisable in Prophet.
    
ii) All the tests are stored in the Results.xlsx file
  