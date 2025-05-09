# PivotTable GetData Method

## Business Description
Returns the value for the a data filed in a PivotTable.

## Behavior
Returns the value for the a data filed in a PivotTable.

## Example Usage
```vba
Msgbox ActiveSheet.PivotTables(1) _ 
 .GetData("'Sum of Revenue' Apples January")
```