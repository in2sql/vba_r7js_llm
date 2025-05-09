# SparklineGroup Object

## Business Description
Represents a group of sparklines.

## Behavior
Represents a group of sparklines.

## Example Usage
```vba
Dim mySG As SparklineGroup 
Set mySG = Range("$A$1:$A$4").SparklineGroups.Add(Type:=xlSparkColumn, SourceData:= _ 
 "Sheet2!B1:E4") 
 
mySG.SeriesColor.Color = RGB(255, 0, 0)
```