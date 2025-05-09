# Workbook BeforePrint Event

## Business Description
Occurs before the workbook (or anything in it) is printed.

## Behavior
Occurs before the workbook (or anything in it) is printed.

## Example Usage
```vba
Private Sub Workbook_BeforePrint(Cancel As Boolean) 
 For Each wk in Worksheets 
 wk.Calculate 
 Next 
End Sub
```