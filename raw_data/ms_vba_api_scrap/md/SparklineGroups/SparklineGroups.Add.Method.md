# SparklineGroups Add Method

## Business Description
Creates a new sparkline group and returns a SparklineGroup object.

## Behavior
Creates a new sparkline group and returns aSparklineGroupobject.

## Example Usage
```vba
Range("$A$1:$A$4").SparklineGroups.Add Type:=xlSparkColumn, SourceData:= _ 
 "Sheet2!B1:E4"
```