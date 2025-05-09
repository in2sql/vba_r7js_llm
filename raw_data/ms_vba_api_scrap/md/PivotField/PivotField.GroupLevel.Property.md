# PivotField GroupLevel Property

## Business Description
Returns the placement of the specified field within a group of fields (if the field is a member of a grouped set of fields). Read-only.

## Behavior
Returns the placement of the specified field within a group of fields (if the field is a member of a grouped set of fields). Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
If ActiveCell.PivotField.GroupLevel= 1 Then 
 MsgBox "This is the highest-level parent field." 
End If
```