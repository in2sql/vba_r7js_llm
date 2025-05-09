# Workbook CheckIn Method

## Business Description
Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.

## Behavior
Returns a workbook from a local computer to a server, and sets the local workbook to read-only so that it cannot be edited locally. Calling this method will also close the workbook.

## Example Usage
```vba
Sub CheckInOut(strWkbCheckIn As String) 
 
 ' Determine if workbook can be checked in. 
 If Workbooks(strWkbCheckIn).CanCheckIn = True Then 
 Workbooks(strWkbCheckIn).CheckInMsgBox strWkbCheckIn & " has been checked in." 
 Else 
 MsgBox "This file cannot be checked in " & _ 
 "at this time. Please try again later." 
 End If 
 
End Sub
```