# PivotTable ShowPages Method

## Business Description
Creates a new PivotTable report for each item in the page field. Each new report is created on a new worksheet.

## Behavior
Creates a new PivotTable report for each item in the page field. Each new report is created on a new worksheet.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.ShowPages"Country"
```