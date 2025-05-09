# Chart SetSourceData Method

## Business Description
Sets the source data range for the chart.

## Behavior
Sets the source data range for the chart.

## Example Usage
```vba
Charts(1).SetSourceDataSource:=Sheets(1).Range("a1:a10"), _ 
 PlotBy:=xlColumns
```