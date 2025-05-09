# Series Values Property

## Business Description
Returns or sets a Variant value that represents a collection of all the values in the series.

## Behavior
Returns or sets aVariantvalue that represents a collection of all the values in the series.

## Example Usage
```vba
Charts("Chart1").SeriesCollection(1).Values = _ 
 Worksheets("Sheet1").Range("C5:T5")
```