# Workbook UserStatus Property

## Business Description
Returns a 1-based, two-dimensional array that provides information about each user who has the workbook open as a shared list. Read-only Variant.

## Behavior
Returns a 1-based, two-dimensional array that provides information about each user who has the workbook open as a shared list.  Read-onlyVariant.

## Example Usage
```vba
users = ActiveWorkbook.UserStatusWith Workbooks.Add.Sheets(1) 
 For row = 1 To UBound(users, 1) 
 .Cells(row, 1) = users(row, 1) 
 .Cells(row, 2) = users(row, 2) 
 Select Case users(row, 3) 
 Case 1 
 .Cells(row, 3).Value = "Exclusive" 
 Case 2 
 .Cells(row, 3).Value = "Shared" 
 End Select 
 Next 
End With
```