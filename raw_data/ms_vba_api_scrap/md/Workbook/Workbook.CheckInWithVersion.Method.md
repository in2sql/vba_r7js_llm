# Workbook CheckInWithVersion Method

## Business Description
Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally.

## Behavior
Saves a workbook to a server from a local computer, and sets the local workbook to read-only so that it cannot be edited locally.

## Example Usage
```vba
Private Sub WorkbookCheckIn() 
 If ActiveWorkbook.CanCheckIn Then 
 ActiveWorkbook.CheckInWithVersion _ 
 True, _ 
 "My updates.", _ 
 True, _ 
 XlCheckInVersionType.xlCheckInMinorVersion 
 Else 
 MessageBox.Show ("This workbook cannot be checked in") 
 End If 
End Sub
```