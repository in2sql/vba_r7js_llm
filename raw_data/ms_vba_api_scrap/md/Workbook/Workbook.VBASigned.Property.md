# Workbook VBASigned Property

## Business Description
True if the Visual Basic for Applications project for the specified workbook has been digitally signed. Read-only Boolean.

## Behavior
Trueif the Visual Basic for Applications project for the specified workbook has been digitally signed. Read-onlyBoolean.

## Example Usage
```vba
Workbooks.Open FileName:="c:\My Documents\mybook.xls", _ 
 ReadOnly:=False 
If Workbook.VBASigned= False Then 
 MsgBox "Warning! The project " _ & 
 "has not been digitally signed." _ & 
 , vbCritical, "Digital Signature Warning" 
End If
```