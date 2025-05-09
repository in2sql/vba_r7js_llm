# Scenarios CreateSummary Method

## Business Description
Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet. Variant.

## Behavior
Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet.Variant.

## Example Usage
```vba
Worksheets("Sheet1").Scenarios.CreateSummary_ 
 ResultCells := Worksheets("Sheet1").Range("C4:C9")
```