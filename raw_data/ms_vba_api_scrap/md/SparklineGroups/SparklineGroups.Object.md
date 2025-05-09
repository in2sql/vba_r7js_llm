# SparklineGroups Object

## Business Description
Represents a collection of sparkline groups.

## Behavior
Represents a collection of sparkline groups.

## Example Usage
```vba
Range("A1:A4").Select 
Selection.SparklineGroups.Group Location := Range("A1") 
Selection.SparklineGroups.Item(1).Points.Markers.Visible = True 
Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = 255
```