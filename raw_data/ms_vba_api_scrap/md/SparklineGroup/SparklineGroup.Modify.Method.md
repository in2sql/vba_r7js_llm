# SparklineGroup Modify Method

## Business Description
Sets the location and the source data for the sparkline group.

## Behavior
Sets the location and the source data for the sparkline group.

## Example Usage
```vba
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).Modify Location:=Range("$A$1:$A$3"), SourceData:="Sheet1!B1:D3"
```