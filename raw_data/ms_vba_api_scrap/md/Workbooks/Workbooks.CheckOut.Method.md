# Workbooks CheckOut Method

## Business Description
Returns a String representing a specified workbook from a server to a local computer for editing.

## Behavior
Returns aStringrepresenting a specified workbook from a server to a local computer for editing.

## Example Usage
```vba
Sub UseCheckOut(docCheckOut As String) 
 
 ' Determine if workbook can be checked out. 
 If Workbooks.CanCheckOut(docCheckOut) = True Then 
 Workbooks.CheckOutdocCheckOut 
 Else 
 MsgBox "Unable to check out this document at this time." 
 End If 
 
End Sub
```