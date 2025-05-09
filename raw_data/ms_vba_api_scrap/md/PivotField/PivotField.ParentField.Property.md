# PivotField ParentField Property

## Business Description
Returns a PivotField object that represents the PivotTable field that's the group parent of the specified object. The field must be grouped and must have a parent field. Read-only.

## Behavior
Returns aPivotFieldobject that represents the PivotTable field that's the group parent of the specified object. The field must be grouped and must have a parent field. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
MsgBox "The active field is a child of the field " & _ 
 ActiveCell.PivotField.ParentField.Name
```