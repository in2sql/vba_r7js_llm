# SparklineGroup ModifySourceData Method

## Business Description
Sets the range that represents the source data for the sparkline group.

## Behavior
Sets the range that represents the source data for the sparkline group.

## Example Usage
```vba
Range("A1:A4").Select 
ActiveCell.SparklineGroups.Item(1).ModifySourceData "B1:D4"
```