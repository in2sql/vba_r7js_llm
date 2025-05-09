# Range DataSeries Method

## Business Description
Creates a data series in the specified range. Variant.

## Behavior
Creates a data series in the specified range.Variant.

## Example Usage
```vba
Set dateRange = Worksheets("Sheet1").Range("A1:A12") 
Worksheets("Sheet1").Range("A1").Formula = "31-JAN-1996" 
dateRange.DataSeriesType:=xlChronological, Date:=xlMonth
```