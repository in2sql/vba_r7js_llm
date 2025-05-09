# Range LocationInTable Property

## Business Description
Returns a constant that describes the part of the PivotTable report that contains the upper-left corner of the specified range. Can be one of the following XlLocationInTable. constants. Read-only Long.

## Behavior
Returns a constant that describes the part of thePivotTablereport that contains the upper-left corner of the specified range. Can be one of the followingXlLocationInTable. constants. Read-onlyLong.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Select Case ActiveCell.LocationInTableCase Is = xlRowHeader 
 MsgBox "Active cell is part of a row header" 
Case Is = xlColumnHeader 
 MsgBox "Active cell is part of a column header" 
Case Is = xlPageHeader 
 MsgBox "Active cell is part of a page header" 
Case Is = xlDataHeader 
 MsgBox "Active cell is part of a data header" 
Case Is = xlRowItem 
 MsgBox "Active cell is part of a row item" 
Case Is = xlColumnItem 
 MsgBox "Active cell is part of a column item" 
Case Is = xlPageItem 
 MsgBox "Active cell is part of a page item" 
Case Is = xlDataItem 
 MsgBox "Active cell is part of a data item" 
Case Is = xlTableBody 
 MsgBox "Active cell is part of the table body" 
End Select
```