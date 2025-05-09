# Range CurrentRegion Property

## Business Description
Returns a Range object that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns. Read-only.

## Behavior
Returns aRangeobject that represents the current region. The current region is a range bounded by any combination of blank rows and blank columns. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveCell.CurrentRegion.Select
```