# Workbook NewChart Event

## Business Description
Occurs when a new chart is created in the workbook.

## Behavior
Occurs when a new chart is created in the workbook.

## Example Usage
```vba
Private Sub Workbook_NewChart(ByVal Ch As Chart) 
 MsgBox ("A new chart of the following chart type was added: " & Ch.ChartType) 
End Sub
```