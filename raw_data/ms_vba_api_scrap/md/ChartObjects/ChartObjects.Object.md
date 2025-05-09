# ChartObjects Object

## Business Description
A collection of all the ChartObject objects on the specified chart sheet, dialog sheet, or worksheet.

## Behavior
A collection of all theChartObjectobjects on the specified chart sheet, dialog sheet, or worksheet.

## Example Usage
```vba
Dim ch As ChartObject 
Set ch = Worksheets("sheet1").ChartObjects.Add(100, 30, 400, 250) 
ch.Chart.ChartWizard source:=Worksheets("sheet1").Range("a1:a20"), _ 
 gallery:=xlLine, title:="New Chart"
```