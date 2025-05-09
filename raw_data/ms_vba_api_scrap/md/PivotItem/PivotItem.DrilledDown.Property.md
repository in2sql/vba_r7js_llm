# PivotItem DrilledDown Property

## Business Description
True if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/write Boolean.

## Behavior
Trueif the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/writeBoolean.

## Example Usage
```vba
ActiveSheet.PivotTables("PivotTable3") _ 
 .PivotFields("state").DrilledDown= False
```