# Workbook KeepChangeHistory Property

## Business Description
True if change tracking is enabled for the shared workbook. Read/write Boolean.

## Behavior
Trueif change tracking is enabled for the shared workbook. Read/writeBoolean.

## Example Usage
```vba
With ActiveWorkbook 
 If .KeepChangeHistoryThen 
 .ChangeHistoryDuration = 7 
 End If 
End With
```