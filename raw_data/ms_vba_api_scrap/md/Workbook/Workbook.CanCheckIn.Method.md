# Workbook CanCheckIn Method

## Business Description
True if Microsoft Excel can check in a specified workbook to a server. Read/write Boolean.

## Behavior
Trueif Microsoft Excel can check in a specified workbook to a server. Read/writeBoolean.

## Example Usage
```vba
Sub CheckInOut(strWkbCheckIn As String) 
 
 ' Determine if workbook can be checked in. 
 If Workbooks(strWkbCheckIn).CanCheckIn= True Then 
 Workbooks(strWkbCheckIn).CheckIn 
 MsgBox strWkbCheckIn & " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " & _ 
 "at this time. Please try again later." 
 End If 
 
End Sub
```