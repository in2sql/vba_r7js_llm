# Workbook ShowConflictHistory Property

## Business Description
True if the Conflict History worksheet is visible in the workbook that's open as a shared list. Read/write Boolean.

## Behavior
Trueif the Conflict History worksheet is visible in the workbook that's open as a shared list. Read/writeBoolean.

## Example Usage
```vba
If ActiveWorkbook.MultiUserEditing Then 
 ActiveWorkbook.ShowConflictHistory= True 
End If
```