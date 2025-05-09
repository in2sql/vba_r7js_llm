# Workbook WriteReservedBy Property

## Business Description
Returns the name of the user who currently has write permission for the workbook. Read-only String.

## Behavior
Returns the name of the user who currently has write permission for the workbook. Read-onlyString.

## Example Usage
```vba
With ActiveWorkbook 
 If .WriteReserved = True Then 
 MsgBox "Please contact " & .WriteReservedBy& Chr(13) & _ 
 " if you need to insert data in this workbook." 
 End If 
End With
```