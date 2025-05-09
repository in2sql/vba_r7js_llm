# PivotItem Position Property

## Business Description
Returns or sets a Long value that represents the position of the item in its field, if the item is currently showing.

## Behavior
Returns or sets aLongvalue that represents the position of the item in its field, if the item is currently showing.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
MsgBox "The active item is in position number " & _ 
 ActiveCell.PivotItem.Position
```