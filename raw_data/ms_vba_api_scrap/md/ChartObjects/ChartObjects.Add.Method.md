# ChartObjects Add Method

## Business Description
Creates a new embedded chart.

## Behavior
Creates a new embedded chart.

## Example Usage
```vba
Set co = Sheets("Sheet1").ChartObjects.Add(50, 40, 200, 100) 
co.Chart.ChartWizard Source:=Worksheets("Sheet1").Range("A1:B2"), _ 
 Gallery:=xlColumn, Format:=6, PlotBy:=xlColumns, _ 
 CategoryLabels:=1, SeriesLabels:=0, HasLegend:=1
```