# Workbook SheetCalculate Event

## Business Description
Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.

## Behavior
Occurs after any worksheet is recalculated or after any changed data is plotted on a chart.

## Example Usage
```vba
Private Sub Workbook_SheetCalculate(ByVal Sh As Object) 
 With Worksheets(1) 
 .Range("a1:a100").Sort Key1:=.Range("a1") 
 End With 
End Sub
```