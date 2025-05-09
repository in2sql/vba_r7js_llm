# Scenarios Add Method

## Business Description
Creates a new scenario and adds it to the list of scenarios that are available for the current worksheet.

## Behavior
Creates a new scenario and adds it to the list of scenarios that are available for the current worksheet.

## Example Usage
```vba
Worksheets("Sheet1").Scenarios.AddName:="Best Case", _ 
 ChangingCells:=Worksheets("Sheet1").Range("A1:A4"), _ 
 Values:=Array(23, 5, 6, 21), _ 
 Comment:="Most favorable outcome."
```