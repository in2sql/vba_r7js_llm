# Workbook ProtectStructure Property

## Business Description
True if the order of the sheets in the workbook is protected. Read-only Boolean.

## Behavior
Trueif the order of the sheets in the workbook is protected. Read-onlyBoolean.

## Example Usage
```vba
If ActiveWorkbook.ProtectStructure= True Then 
 MsgBox "Remember, you cannot delete, add, or change " & _ 
 Chr(13) & _ 
 "the location of any sheets in this workbook." 
End If
```