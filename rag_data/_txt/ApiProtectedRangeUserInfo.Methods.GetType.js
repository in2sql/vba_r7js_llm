**Description / Описание**

This code example changes the user-protected range in a worksheet.  
Этот пример изменяет защищенный диапазон пользователя на листе.

```javascript
// This example changes the user protected range.
// Этот пример изменяет защищенный диапазон пользователя.

var oWorksheet = Api.GetActiveSheet();
// Get the active worksheet.

oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");
// Add a protected range named "protectedRange" to cells A1:B1 and add a user with view permissions.

var protectedRange = oWorksheet.GetProtectedRange("protectedRange");
// Retrieve the protected range named "protectedRange".

var userInfo = protectedRange.GetUser("userId");
// Get information about the user with ID "userId".

var userType = userInfo.GetType();
// Get the type of the user.

oWorksheet.GetRange("A3").SetValue("Type: " + userType); 
// Set the value of cell A3 to display the user type.
```

```vba
' This code example changes the user protected range.
' Этот пример изменяет защищенный диапазон пользователя.

Sub ChangeUserProtectedRange()
    ' Get the active worksheet.
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a protected range named "protectedRange" to cells A1:B1 and add a user with view permissions.
    oWorksheet.API.AddProtectedRange "protectedRange", "$A$1:$B$1"
    oWorksheet.API.GetProtectedRange("protectedRange").AddUser "userId", "name", "CanView"
    
    ' Retrieve the protected range named "protectedRange".
    Dim protectedRange As Object
    Set protectedRange = oWorksheet.API.GetProtectedRange("protectedRange")
    
    ' Get information about the user with ID "userId".
    Dim userInfo As Object
    Set userInfo = protectedRange.GetUser("userId")
    
    ' Get the type of the user.
    Dim userType As String
    userType = userInfo.GetType()
    
    ' Set the value of cell A3 to display the user type.
    oWorksheet.Range("A3").Value = "Type: " & userType
End Sub
```