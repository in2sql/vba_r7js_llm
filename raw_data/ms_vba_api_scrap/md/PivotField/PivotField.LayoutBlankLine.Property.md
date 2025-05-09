# PivotField LayoutBlankLine Property

## Business Description
True if a blank row is inserted after the specified row field in a PivotTable report. The default value is False. Read/write Boolean.

## Behavior
Trueif a blank row is inserted after the specified row field in a PivotTable report. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutBlankLine= True 
End With
```