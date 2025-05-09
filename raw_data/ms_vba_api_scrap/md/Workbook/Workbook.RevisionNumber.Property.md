# Workbook RevisionNumber Property

## Business Description
Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero). Read-only Long.

## Behavior
Returns the number of times the workbook has been saved while open as a shared list. If the workbook is open in exclusive mode, this property returns 0 (zero). Read-onlyLong.

## Example Usage
```vba
If ActiveWorkbook.RevisionNumber = 0 Then 
 ActiveWorkbook.SaveAs _ 
 filename:=ActiveWorkbook.FullName, _ 
 accessMode:=xlShared, _ 
 conflictResolution:= _ 
 xlOtherSessionChanges 
End If
```