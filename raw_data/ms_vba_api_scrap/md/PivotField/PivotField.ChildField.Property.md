# PivotField ChildField Property

## Business Description
Returns a PivotField object that represents the child field for the specified field (if the field is grouped and has a child field). Read-only.

## Behavior
Returns aPivotFieldobject that represents the child field for the specified field (if the field is grouped and has a child field). Read-only.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox "The name of the child field is " & _ 
 pvtTable.PivotFields("REGION2").ChildField.Name
```