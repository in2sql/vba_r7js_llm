# PivotItem ParentItem Property

## Business Description
Returns a PivotItem object that represents the parent PivotTable item in the parent PivotField object (the field must be grouped so that it has a parent). Read-only.

## Behavior
Returns aPivotItemobject that represents the parent PivotTable item in the parentPivotFieldobject (the field must be grouped so that it has a parent). Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
MsgBox "This item is a subitem of " & _ 
 ActiveCell.PivotItem.ParentItem.Name
```