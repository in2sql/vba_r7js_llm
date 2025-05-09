# Workbook CreateBackup Property

## Business Description
True if a backup file is created when this file is saved. Read-only Boolean.

## Behavior
Trueif a backup file is created when this file is saved. Read-onlyBoolean.

## Example Usage
```vba
If ActiveWorkbook.CreateBackup= True Then 
 MsgBox "Remember, there is a backup copy of this workbook" 
End If
```