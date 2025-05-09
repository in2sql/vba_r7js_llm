# SparklineGroup ModifyLocation Method

## Business Description
Sets the associated Range object to modify the location of the sparkline group.

## Behavior
Sets the associatedRangeobject to modify the location of the sparkline group.

## Example Usage
```vba
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).ModifyLocation Range("$A$10:$A$14")
```