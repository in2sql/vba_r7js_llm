# PivotCell Range Property

## Business Description
Returns a Range object that represents the range the specified PivotCell applies to.

## Behavior
Returns aRangeobject that represents the range the specified PivotCell applies to.

## Example Usage
```vba
Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).RangeActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column
```