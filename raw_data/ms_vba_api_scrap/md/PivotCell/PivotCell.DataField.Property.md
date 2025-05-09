# PivotCell DataField Property

## Business Description
Returns a PivotField object that corresponds to the selected data field.

## Behavior
Returns aPivotFieldobject that corresponds to the selected data field.

## Example Usage
```vba
Sub CheckDataField() 
 
 On Error GoTo Not_In_DataField 
 
 MsgBox Application.Range("L10").PivotCell.DataFieldExit Sub 
 
Not_In_DataField: 
 MsgBox "The selected range is not in the data field of the PivotTable." 
 
End Sub
```