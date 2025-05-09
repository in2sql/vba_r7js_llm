# ColorScale Object

## Business Description
Represents a color scale conditional formatting rule.

## Behavior
Represents a color scale conditional formatting rule.

## Example Usage
```vba
Sub CreateColorScaleCF() 
 
 Dim cfColorScale AsColorScale'Fill cells with sample data from 1 to 10 
 With ActiveSheet 
 .Range("C1") = 1 
 .Range("C2") = 2 
 .Range("C1:C2").AutoFill Destination:=Range("C1:C10") 
 End With 
 
 Range("C1:C10").Select 
 
 'Create a two-color ColorScale object for the created sample data range 
 Set cfColorScale = Selection.FormatConditions.AddColorScale(ColorScaleType:=2) 
 
 'Set the minimum threshold to red and maximum threshold to blue 
 cfColorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0) 
 cfColorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255) 
 
End Sub
```