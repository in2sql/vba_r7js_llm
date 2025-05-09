# Workbook WriteReserved Property

## Business Description
True if the workbook is write-reserved. Read-only Boolean.

## Behavior
Trueif the workbook is write-reserved. Read-onlyBoolean.

## Example Usage
```vba
With ActiveWorkbook 
 If .WriteReserved= True Then 
 MsgBox "Please contact " & .WriteReservedBy & Chr(13) & _ 
 " if you need to insert data in this workbook." 
 End If 
End With
```