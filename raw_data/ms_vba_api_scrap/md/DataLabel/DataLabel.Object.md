# DataLabel Object

## Business Description
Represents the data label on a chart point or trendline.

## Behavior
Represents the data label on a chart point or trendline.

## Example Usage
```vba
With Charts("chart1").SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = False 
 .DisplayEquation = True 
 Worksheets("sheet1").Range("a1").Value = .DataLabel.Text 
End With
```