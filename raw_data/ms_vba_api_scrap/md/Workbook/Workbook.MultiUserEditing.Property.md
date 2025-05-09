# Workbook MultiUserEditing Property

## Business Description
True if the workbook is open as a shared list. Read-only Boolean.

## Behavior
Trueif the workbook is open as a shared list. Read-onlyBoolean.

## Example Usage
```vba
If Not ActiveWorkbook.MultiUserEditingThen 
 ActiveWorkbook.SaveAs fileName:=ActiveWorkbook.FullName, _ 
 accessMode:=xlShared 
End If
```