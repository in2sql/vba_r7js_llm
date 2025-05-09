# SeriesCollection Object

## Business Description
A collection of all the Series objects in the specified chart or chart group.

## Behavior
A collection of all theSeriesobjects in the specified chart or chart group.

## Example Usage
```vba
Worksheets(1).ChartObjects(1).Chart. _ 
 SeriesCollection.Extend Worksheets(1).Range("c1:c10")
```