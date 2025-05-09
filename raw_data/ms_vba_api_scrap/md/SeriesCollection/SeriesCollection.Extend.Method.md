# SeriesCollection Extend Method

## Business Description
Adds new data points to an existing series collection.

## Behavior
Adds new data points to an existing series collection.

## Example Usage
```vba
Charts("Chart1").SeriesCollection.Extend_ 
        Source:=Worksheets("Sheet1").Range("B1:B6")
```