# PivotCell PivotField Property

## Business Description
Returns a PivotField object that represents the PivotTable field containing the upper-left corner of the specified range.

## Behavior
Returns aPivotFieldobject that represents the PivotTable field containing the upper-left corner of the specified range.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
MsgBox "The active cell is in the field " & _ 
 ActiveCell.PivotField.Name
```