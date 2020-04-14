# MISI-Calculator
MISI Calculator - MATLAB code for GUI to allow for the calculation of skeletal muscle specific insulin sensitivity from oral glucose tolerance test data. 

The MISI Calculator is a user friendly GUI written in MATLAB (2014b) that allows for the computation of the muscle insulin sensitivity index on user supplied data. Glucose and insulin oral glucose tolerance test data can be uploaded in .xlsx (excel workbook) format. Sampling time points and measurement units can also be specified. The calculator computes MISI using the standard method proposed by Abdul Ghani et al. (2007) and also allows for the computation of MISI using the modified cubic spline method, which is recommended with five or less time points or with unequal sampling frequency in the OGTT data (O'Donovan et al. 2019). The calculator also allows the user to flag glucose curves which may produce erroneous MISI values for manual inspection. Finally, the user may choose to display the glucose and insulin curves of flagged individuals or save them to folder for later inspection.

The MISI Calculator can also be installed as a stand-alone app for windows and mac using MATLAB Runtime from https://www.maastrichtuniversity.nl/macsbio-misi-calculator




