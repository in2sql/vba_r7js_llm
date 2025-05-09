# Workbook HasPassword Property

## Business Description
True if the workbook has a protection password. Read-only Boolean.

## Behavior
Trueif the workbook has a protection password. Read-onlyBoolean.

## Example Usage
```vba
If ActiveWorkbook.HasPassword= True Then 
 MsgBox "Remember to obtain the workbook password" & Chr(13) & _ 
 " from the Network Administrator." 
End If
```