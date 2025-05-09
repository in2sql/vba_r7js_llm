# PivotField TotalLevels Property

## Business Description
Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based, TotalLevels returns the value 1. Read-only Long.

## Behavior
Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based,TotalLevelsreturns the value 1. Read-onlyLong.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
MsgBox "This group has " & _ 
 ActiveCell.PivotField.TotalLevels& " levels
```