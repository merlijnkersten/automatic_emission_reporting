# Automatic emission reporting
_April-June 2022, merlijn_


### Goal
Automatically populate cells in `annex_iV_rev2022_v1.xlsx` (file A) and `GovReg_Proj_T1a_T1b_T5a_T5b_v1.1.xlsx` (file B) with data from VEDA/TIMES-CZ.


### Method 
1. For every pollutant, find the sum of PV values for every NFR code, given an attribute, scenario and period. 
2. Save the sums to a `vector` and write this to one of the Excel files.

### Run code
Check that all the parameters in the code are set correctly. They are labelled by a `PARAMETER` comment. This includes ensuring that the order of the NFR codes of `rows_a.csv` (for file A) and `rows_b.csv` (for file B) matches the NFR codes in columns `B` and `A` of their respective Excel files.

The `GovReg_Proj_T1a_T1b_T5a_T5b_v1.1.xlsx` needs to be edited before use (it causes the code to break). I copied the values of a sheet (e.g. `Table1A`) to a new Excel spreadsheet, and then copied the layout. This results in a sheet similar to the original but without any of the functions and equations (see `GovReg_Proj_T1a_T1b_T5a_T5b_v1.1 - Table1a.xlsx` ). 

### Future improvements
The code needs to be tested further, and some of its functions could be rewritten to be more "`R`-native".