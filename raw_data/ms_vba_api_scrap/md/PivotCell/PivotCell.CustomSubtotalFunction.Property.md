# PivotCell CustomSubtotalFunction Property

## Business Description
Returns the custom subtotal function field setting of a PivotCell object. Read-only XlConsolidationFunction.

## Behavior
Returns the custom subtotal function field setting of aPivotCellobject. Read-onlyXlConsolidationFunction.

## Example Usage
```vba
Sub UseCustomSubtotalFunction() 
 
 On Error GoTo Not_A_Function 
 
 ' Determine if custom subtotal function is a count function. 
 If Application.Range("C20").PivotCell.CustomSubtotalFunction= xlCount Then 
 MsgBox "The custom subtotal function is a Count." 
 Else 
 MsgBox "The custom subtotal function is not a Count." 
 End If 
 Exit Sub 
 
Not_A_Function: 
 MsgBox "The selected cell is not a custom subtotal function." 
 
End Sub
```