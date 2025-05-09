# PivotCell PivotItem Property

## Business Description
Returns a PivotItem object that represents the PivotTable item containing the upper-left corner of the specified range.

## Behavior
Returns aPivotItemobject that represents the PivotTable item containing the upper-left corner of the specified range.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the item " & _ 
 ActiveCell.PivotItem.Name
```