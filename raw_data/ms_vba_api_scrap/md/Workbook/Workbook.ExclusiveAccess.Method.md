# Workbook ExclusiveAccess Method

## Business Description
Assigns the current user exclusive access to the workbook that's open as a shared list.

## Behavior
Assigns the current user exclusive access to the workbook that's open as a shared list.

## Example Usage
```vba
If ActiveWorkbook.MultiUserEditing Then 
 ActiveWorkbook.ExclusiveAccessEnd If
```