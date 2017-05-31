# B-LacBarcoding_08-21-2009

Date Written: 08/21/2009

Industry: Medical Device Manufacturer

Device: Blood Analyzer

Market: Human

Platform: B-Lac Sensor Panel – pH, CO2 (torr), O2 (torr), & Lactate (mM)

GUI:
GUI.png

Sample Raw Data Files:
“DB1716.805”,  “DB1718.805”, “DB2557.805”, & “DB9223.805”  There are tab delimited text files 

Sample Output:
SampleOutput_BarcodeResults.png & SampleOutput_StasticalAnalysis.png

Application Description:

QC technicians measure the concentration of pH, CO2 (torr), O2 (torr), & Lactate (mM) within 5 standard buffers using a given lot of B-Lac sensors.  The measurements are done in quadruplicate across four analyzers.  Three different tonometerized blood samples and a raw blood sample are also sampled across the 4 analyzers.  A reference instrument is utilized for measurement verification.  After data acquisition the instrument database is exported as tab delimited text file.

This application mines the appropriate spectroscopic data so that the calibration constants can be calculated using the data from the 5 buffer standards.  These constants are then used to determine the concentrations of each analyte in each of the aforementioned blood samples.  Once calculated the values are compared to the known theoretical values and the coefficient of variation (CV) is determined.  If the CV values are within acceptable limits as set by the FDA and CLIA guidelines the calibration constants are used to generate a barcode for the corresponding lot of sensors.
