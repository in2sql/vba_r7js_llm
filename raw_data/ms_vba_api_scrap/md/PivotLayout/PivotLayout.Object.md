# PivotLayout Object

## Business Description
Represents the placement of fields in a PivotChart report.

## Behavior
Represents the placement of fields in a PivotChart report.

## Example Usage
```vba
Sub ListFieldNames 
 
 Dim objNewSheet As Worksheet 
 Dim intRow As Integer 
 Dim objPF As PivotField 
 
 Set objNewSheet = Worksheets.Add 
 
 intRow = 1 
 
 For Each objPF In _ 
 Charts("Chart1").PivotLayout.PivotFields 
 
 objNewSheet.Cells(intRow, 1).Value = objPF.Caption 
 
 intRow = intRow + 1 
 
 Next objPF 
 
End Sub
```