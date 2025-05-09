# Series XValues Property

## Business Description
Returns or sets an array of x values for a chart series. The XValues property can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/write Variant.

## Behavior
Returns or sets an array of x values for a chart series. TheXValuesproperty can be set to a range on a worksheet or to an array of values, but it cannot be a combination of both. Read/writeVariant.

## Example Usage
```vba
Charts("Chart1").SeriesCollection(1).XValues= _ 
 Worksheets("Sheet1").Range("B1:B5")
```