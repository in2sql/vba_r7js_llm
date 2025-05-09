# Charts Object

## Business Description
A collection of all the chart sheets in the specified or active workbook.

## Behavior
A collection of all the chart sheets in the specified or active workbook.

## Example Usage
```vba
WithCharts.Add 
 .ChartWizard source:=Worksheets("Sheet1").Range("A1:A20"), _ 
 Gallery:=xlLine, Title:="February Data" 
End With
```