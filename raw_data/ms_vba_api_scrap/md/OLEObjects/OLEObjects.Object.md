# OLEObjects Object

## Business Description
A collection of all the OLEObject objects on the specified worksheet.

## Behavior
A collection of all theOLEObjectobjects on the specified worksheet.

## Example Usage
```vba
Private Sub chkFinished_Click() 
 ActiveSheet.OLEObjects("CheckBox1").Object.Value = 1 
End Sub
```