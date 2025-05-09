# ColorScaleCriteria Object

## Business Description
A collection of ColorScaleCriterion objects that represents all of the criteria for a color scale conditional format. Each criterion specifies the minimum, midpoint, or maximum threshold for the color scale.

## Behavior
A collection ofColorScaleCriterionobjects that represents all of the criteria for a color scale conditional format. Each criterion specifies the minimum, midpoint, or maximum threshold for the color scale.

## Example Usage
```vba
Sub CreateColorScaleCF() 
 
 Dim cfColorScale As ColorScale 
 
 'Fill cells with sample data from 1 to 10 
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