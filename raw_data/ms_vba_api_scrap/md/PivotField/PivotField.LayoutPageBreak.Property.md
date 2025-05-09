# PivotField LayoutPageBreak Property

## Business Description
True if a page break is inserted after each field. The default value is False. Read/write Boolean.

## Behavior
Trueif a page break is inserted after each field. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
With ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("state") 
 .LayoutPageBreak= True 
End With
```