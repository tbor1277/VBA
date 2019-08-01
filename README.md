# Visual Basic
This is a collection of VBA Scripts that I use.


# Excel
## [extract_value](Excel/extract_value.bas)
This script extracts the value inside a parenthesis "()"

### How to Use
Enter formula command:
=extract_value()

## [PVE](Excel/PVE.bas)
PVE stands for parenthesis value extractor that extracts the data inside parenthesis "()".

This is my Macro+Script combo. The key strokes of Ctrl+u, Ctrl+i, Ctrl+o will execute this.
Keep an eye out to the target files. ie. Macro.xlsm

1. copycolumns
-The script copies data in column D to a Workbook named "Macro.xlsm"(in cell A1),

2. extract Macro(Ctrl+i)
-extracts the data inside parenthesis "()" and paste it another cell.

3. secondcopy Macro(Ctrl+o)
-copy column R to a Workbook named "Macro.xlsm"(in cell A1)

## [OpenFolder](Excel/OpenFolder.bas)
Opens a target folder using Shell command-explorer.exe. Gives out an error message if the item/path/file does not exists or is incorrect.

## [DoesFileExist](Excel/DoesFileExist.bas)
Determines if a target file exists or not. Returns boolean values `"Exists"` or `"Not Found"`.

## [DeleteFile](Excel/DeleteFile.bas)
Deletes a target file. Gives out an error message if the item/path/file does not exists or is incorrect. **Use with caution!**

## [CreatePowerPoint](Excel/CreatePowerPoint.bas)
Pulls all graphs in an excel sheet and links it to a powerpoint presentation.
