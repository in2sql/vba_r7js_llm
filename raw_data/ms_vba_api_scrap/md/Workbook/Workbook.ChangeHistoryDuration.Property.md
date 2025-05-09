# Workbook ChangeHistoryDuration Property

## Business Description
Returns or sets the number of days shown in the shared workbook's change history. Read/write Long.

## Behavior
Returns or sets the number of days shown in the shared workbook's change history. Read/writeLong.

## Example Usage
```vba
With ActiveWorkbook 
 If .KeepChangeHistory Then 
 .ChangeHistoryDuration= 7 
 End If 
End With
```