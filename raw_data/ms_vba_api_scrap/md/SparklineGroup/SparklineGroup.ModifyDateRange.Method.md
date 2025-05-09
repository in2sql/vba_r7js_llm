# SparklineGroup ModifyDateRange Method

## Business Description
Sets the date range for the sparkline group.

## Behavior
Sets the date range for the sparkline group.

## Example Usage
```vba
Range("A2:A5").Select 
ActiveCell.SparklineGroups.Item(1).ModifyDateRange "Sheet1!B1:E1"
```