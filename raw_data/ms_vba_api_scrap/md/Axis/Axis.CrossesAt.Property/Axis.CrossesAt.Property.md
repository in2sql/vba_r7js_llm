# Axis.CrossesAt property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub Chart() 
 
 ' Create a sample source of data. 
 Range("A1") = "2" 
 Range("A2") = "4" 
 Range("A3") = "6" 
 Range("A4") = "3" 
 
 ' Create a chart based on the sample source of data. 
 Charts.Add 
 
 With ActiveChart 
 .ChartType = xlLineMarkersStacked 
 .SetSourceData Source:=Sheets("Sheet1").Range("A1:A4"), PlotBy:= xlColumns 
 .Location Where:=xlLocationAsObject, Name:="Sheet1" 
 End With 
 
 ' Set the category axis to cross the value axis at value 3. 
 ActiveChart.Axes(xlValue).Select 
 Selection.CrossesAt = 3 
 
End Sub
```

## Remarks
Setting this property causes the Crosses property to change to xlAxisCrossesCustom.

## Example
```vba
Sub Chart() 
 
 ' Create a sample source of data. 
 Range("A1") = "2" 
 Range("A2") = "4" 
 Range("A3") = "6" 
 Range("A4") = "3" 
 
 ' Create a chart based on the sample source of data. 
 Charts.Add 
 
 With ActiveChart 
 .ChartType = xlLineMarkersStacked 
 .SetSourceData Source:=Sheets("Sheet1").Range("A1:A4"), PlotBy:= xlColumns 
 .Location Where:=xlLocationAsObject, Name:="Sheet1" 
 End With 
 
 ' Set the category axis to cross the value axis at value 3. 
 ActiveChart.Axes(xlValue).Select 
 Selection.CrossesAt = 3 
 
End Sub
```

