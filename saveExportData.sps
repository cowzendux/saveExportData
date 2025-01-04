* Encoding: UTF-8.
* saveExportData
* by Jamie DeCoster

* This function Saves the current dataset to the supplied path as a .sav, Converts it to Excel format,
* Creates a data dictionary and saves it as a .spv, and Converts the data dictionary to word format.
* It also can optionally drop a list of variables from the save file, or keep a list
* of variables in the save file.
    
**** Usage: saveExportData(filepath, keep=None, drop=None)
**** "filepath" is a string indicating the directory and filename for the data set. The extension should 
* not be included. The filepath can make use of file handles.
**** "keep" is a string that will be be preceeded by "/keep " in the save command. Any variables not 
* mentioned in this string will be excluded from the data set.
**** "drop" is a string that will be be preceeded by "/drop " in the save command. Any variables 
* mentioned in this string will be excluded from the data set. If you provide values for both the 
* "keep" and "drop" arguments, the argument for "drop" will be ignored.

**** Example: saveExportData(filepath = "D:/Banking Time/Data/3 Cleaned/Discipline data by StudID",
keep = """StudID RecordedDate
  ExcDis NonExcDis
  DisLog_1 to DisLog_15""")
**** This will take the currently active dataset and save it as 
* D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.sav. It will then create 3 additional files.
* D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.xlsx : Excel conversion of data set
* D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.spv : Data dictionary saved as SPSS output file
* D:/Banking Time/Data/3 Cleaned/Discipline data by StudID.docx : Data dictionary saved as Word doc
**** This file will contain the variables StudID, RecordedDate, ExcDis, NonExcDis, and  the variables 
* between DisLog_1 and DisLog_15. All other variables in the data set will be dropped.

begin program python3.
import spss

def saveExportData(filepath, keep=None, drop=None):
    # Remove .sav at end of file path if it exists
    if (filepath[-4:].upper() == ".SAV"):
        filepath = filepath[:-4]
            
    submitstring = f"SAVE OUTFILE='{filepath}.sav'"
    if (keep != None):
        submitstring += f"\n  /keep {keep}"
    if (keep == None and drop != None):
        submitstring += f"\n /drop {drop}"
    submitstring += f"\n  /COMPRESSED."
    spss.Submit(submitstring)

    submitstring = f"""SAVE TRANSLATE OUTFILE='{filepath}.xlsx'
  /TYPE=XLS
  /VERSION=12
  /MAP
  /FIELDNAMES VALUE=NAMES
  /CELLS=VALUES
  /REPLACE.

output close all.
GET
  FILE='{filepath}.sav'.
dataset name $dataset.

display dictionary.

OUTPUT SAVE 
 OUTFILE='{filepath}.spv'
 LOCK=NO.
 
OUTPUT EXPORT
  /CONTENTS  EXPORT=ALL  LAYERS=PRINTSETTING  MODELVIEWS=PRINTSETTING
  /DOC  DOCUMENTFILE='{filepath}.docx'
     NOTESCAPTIONS=YES  WIDETABLES=WRAP PAGEBREAKS=YES
     PAGESIZE=INCHES(8.5, 11.0)  TOPMARGIN=INCHES(.5)  BOTTOMMARGIN=INCHES(.5)
     LEFTMARGIN=INCHES(.5)  RIGHTMARGIN=INCHES(.5)."""

    spss.Submit(submitstring)
end program python3.

*******
* Version History
*******
* 2024-07-20 Created
* 2024-09-22 Separated saving of data set from creating dictionary
* 2024-10-13 Removed .sav from end of filepath
