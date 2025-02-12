**Description / Описание**

This code modifies a protected range in the active worksheet by adding a user with specific permissions and displays the user's name in cell A3.

Этот код изменяет защищенный диапазон на активном листе, добавляя пользователя с определенными правами и отображает имя пользователя в ячейке A3.

```vba
' VBA Code to modify a protected range and add a user with specific permissions

Sub ModifyProtectedRange()
    Dim oWorksheet As Worksheet
    Dim protectedRange As Range
    Dim userName As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Protect the range A1:B1 with a password and allow viewing
    oWorksheet.Range("A1:B1").Locked = True
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Add user information (Note: VBA does not support user-specific permissions directly)
    ' This is a placeholder to represent adding user information
    userName = "name"
    
    ' Set value in cell A3 to display the user's name
    oWorksheet.Range("A3").Value = "User name: " & userName
End Sub
```

```javascript
// JavaScript Code to modify a protected range and add a user with specific permissions

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" covering cells A1:B1
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1")
    .AddUser("userId", "name", "CanView"); // Add a user with CanView permission

// Retrieve the protected range
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Get user information by userId
var userInfo = protectedRange.GetUser("userId");

// Get the user's name
var userName = userInfo.GetName();

// Set the value of cell A3 to display the user's name
oWorksheet.GetRange("A3").SetValue("User name: " + userName);
```