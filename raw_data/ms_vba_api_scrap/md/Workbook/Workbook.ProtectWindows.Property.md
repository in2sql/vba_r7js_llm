# Workbook ProtectWindows Property

## Business Description
True if the windows of the workbook are protected. Read-only Boolean.

## Behavior
Trueif the windows of the workbook are protected. Read-onlyBoolean.

## Example Usage
```vba
If ActiveWorkbook.ProtectWindows= True Then 
 MsgBox "Remember, you cannot rearrange any" & _ 
 " window in this workbook." 
End If
```