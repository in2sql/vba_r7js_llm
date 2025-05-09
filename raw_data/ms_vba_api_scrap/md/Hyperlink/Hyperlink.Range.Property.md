# Hyperlink Range Property

## Business Description
Returns a Range object that represents the range the specified hyperlink is attached to.

## Behavior
Returns aRangeobject that represents the range the specified hyperlink is attached to.

## Example Usage
```vba
Workbooks(1).Activate 
Set hr = ActiveSheet.Hyperlinks(1).RangeActiveWindow.ScrollRow = hr.Row 
ActiveWindow.ScrollColumn = hr.Column
```