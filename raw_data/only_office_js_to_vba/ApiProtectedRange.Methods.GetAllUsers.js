**Description:**

This script sets values in cells A1, B1, and C1, adds a protected range from A1 to C1, assigns users with specific permissions, retrieves the users, and sets the name of the first user in cell A3.

Этот скрипт устанавливает значения в ячейки A1, B1 и C1, добавляет защищенный диапазон от A1 до C1, назначает пользователей с определенными разрешениями, извлекает пользователей и записывает имя первого пользователя в ячейку A3.

```vba
' VBA Code

Sub ProtectRangeExample()
    Dim oWorksheet As Worksheet
    Dim protectedRange As Range
    Dim users(1 To 2) As String
    
    ' Get the active sheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in A1, B1, C1
    oWorksheet.Range("A1").Value = "1"
    oWorksheet.Range("B1").Value = "2"
    oWorksheet.Range("C1").Value = "3"
    
    ' Protect the range A1:C1 by locking the cells and protecting the sheet
    oWorksheet.Range("A1:C1").Locked = True
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Assign users (Note: Excel VBA does not support user-based permissions directly)
    ' This is a simulation using comments as Excel handles permissions differently
    users(1) = "John Smith" ' CanEdit
    users(2) = "Mark Potato" ' CanView
    
    ' Set the name of the first user in cell A3
    oWorksheet.Range("A3").Value = users(1)
End Sub
```

```javascript
// OnlyOffice JS Code

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Set values in A1, B1, C1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");
oWorksheet.GetRange("C1").SetValue("3");

// Add a protected range from A1 to C1
oWorksheet.AddProtectedRange("Protected range", "$A$1:$C$1");

// Get the protected range
var oProtectedRange = oWorksheet.GetProtectedRange("Protected range");

// Add users with permissions
oProtectedRange.AddUser("uid-1", "John Smith", "CanEdit");
oProtectedRange.AddUser("uid-2", "Mark Potato", "CanView");

// Get all users and set the first user's name in A3
var aUsers = oProtectedRange.GetAllUsers();
oWorksheet.GetRange("A3").SetValue(aUsers[0].GetName());
```