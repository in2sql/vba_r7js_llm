# Description / Описание

**English:**  
This script sets values in cells A1, B1, and C1, adds a protected range covering these cells, assigns users with specific permissions to the protected range, retrieves all users associated with the range, and sets the name of the first user in cell A3.

**Russian:**  
Этот скрипт устанавливает значения в ячейках A1, B1 и C1, добавляет защищенный диапазон, охватывающий эти ячейки, назначает пользователям конкретные права доступа к защищенному диапазону, получает всех пользователей, связанных с диапазоном, и устанавливает имя первого пользователя в ячейке A3.

```vba
' VBA Code to replicate the OnlyOffice JS functionality

Sub ProtectRangeExample()
    Dim oWorksheet As Worksheet
    Dim protectedRange As Range
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in cells A1, B1, C1
    oWorksheet.Range("A1").Value = "1"
    oWorksheet.Range("B1").Value = "2"
    oWorksheet.Range("C1").Value = "3"
    
    ' Protect the range A1:C1
    With oWorksheet
        ' Unlock all cells first
        .Cells.Locked = False
        ' Lock the specific range
        .Range("A1:C1").Locked = True
        ' Protect the worksheet
        .Protect Password:="YourPassword", UserInterfaceOnly:=True
    End With
    
    ' Note: VBA does not support assigning specific user permissions like OnlyOffice.
    ' The following lines simulate adding user names.
    
    ' Add user names to cell A3
    oWorksheet.Range("A3").Value = "John Smith"
    ' To add the second user, you could use another cell, e.g., B3
    oWorksheet.Range("B3").Value = "Mark Potato"
    
    ' Retrieve and display the first user name from A3
    Dim firstUser As String
    firstUser = oWorksheet.Range("A3").Value
    MsgBox "First user: " & firstUser
End Sub
```

```javascript
// OnlyOffice JS Code to set cell values, protect a range, assign users, and retrieve user names

var oWorksheet = Api.GetActiveSheet();

// Set values in cells A1, B1, C1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");
oWorksheet.GetRange("C1").SetValue("3");

// Add a protected range covering A1:C1
oWorksheet.AddProtectedRange("Protected range", "$A$1:$C$1");

// Get the protected range object
var oProtectedRange = oWorksheet.GetProtectedRange("Protected range");

// Assign users with specific permissions
oProtectedRange.AddUser("uid-1", "John Smith", "CanEdit");
oProtectedRange.AddUser("uid-2", "Mark Potato", "CanView");

// Retrieve all users associated with the protected range
var aUsers = oProtectedRange.GetAllUsers();

// Set the name of the first user in cell A3
oWorksheet.GetRange("A3").SetValue(aUsers[0].GetName());
```