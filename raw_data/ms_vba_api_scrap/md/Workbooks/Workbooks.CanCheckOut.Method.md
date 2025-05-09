# Workbooks CanCheckOut Method

## Business Description
True if Microsoft Excel can check out a specified workbook from a server. Read/write Boolean.

## Behavior
Trueif Microsoft Excel can check out a specified workbook from a server. Read/writeBoolean.

## Example Usage
```vba
Sub UseCanCheckOut(docCheckOut As String) 
 
 ' Determine if workbook can be checked out. 
 If Workbooks.CanCheckOut(Filename:=docCheckOut) = True Then 
 Workbooks.CheckOut (Filename:=docCheckOut) 
 Else 
 MsgBox "You are unable to check out this document at this time." 
 End If 
 
End Sub
```