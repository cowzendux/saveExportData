# saveExportData

SPSS Python macro to save a data set, export it to Excel, create a data dictionary, and export the dictionary to Word

This function 
* Saves the current dataset to the supplied path as a .sav file. It also can optionally drop a list of variables from the save file, or keep a list of variables in the save file.
* Converts the data to Excel format
* Creates a data dictionary and saves it as a .spv
* Converts the data dictionary to word format.

## Usage
**saveExportData(filepath, keep=None, drop=None)**
* "filepath" is a string indicating the directory and filename for the data set. The extension should not be included. The filepath can make use of file handles.
* "keep" is a string that will be be preceeded by "/keep " in the save command. Any variables not mentioned in this string will be excluded from the data set.
* "drop" is a string that will be be preceeded by "/drop " in the save command. Any variables mentioned in this string will be excluded from the data set. If you provide values for both the "keep" and "drop" arguments, the argument for "drop" will be ignored.

## Example
**saveExportData(filepath = "D:/Banking Time/Data/3 Cleaned/Discipline data by StudID",  
keep = """StudID RecordedDate  
  ExcDis NonExcDis  
  DisLog_1 to DisLog_15""")
* This will take the currently active dataset and save it as D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.sav. It will then create 3 additional files.
  * D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.xlsx : Excel conversion of data set
  * D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.spv : Data dictionary saved as SPSS output file
  * D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.docx : Data dictionary saved as Word doc
* This file will contain the variables StudID, RecordedDate, ExcDis, NonExcDis, and  the variables between DisLog_1 and DisLog_15. All other variables in the data set will be dropped.
