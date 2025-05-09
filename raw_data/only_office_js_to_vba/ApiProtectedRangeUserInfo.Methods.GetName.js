**Description:**

*English: This example changes the user protected range.*

*Russian: Этот пример изменяет защищенный диапазон пользователя.*

```javascript
// This example changes the user protected range.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1") // Add a protected range named "protectedRange" covering cells A1 to B1
    .AddUser("userId", "name", "CanView"); // Add a user with ID "userId", name "name", and permission "CanView"
var protectedRange = oWorksheet.GetProtectedRange("protectedRange"); // Retrieve the protected range
var userInfo = protectedRange.GetUser("userId"); // Get user information for "userId"
var userName = userInfo.GetName(); // Get the user's name
oWorksheet.GetRange("A3").SetValue("Name: " + userName); // Set the value of cell A3 to display the user's name
```

```vba
' This example changes the user protected range.
Sub ChangeProtectedRange()
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet() ' Get the active worksheet
    
    ' Add a protected range named "protectedRange" covering cells A1 to B1
    oWorksheet.AddProtectedRange "protectedRange", "$A$1:$B$1"
    
    ' Add a user with ID "userId", name "name", and permission "CanView"
    oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser "userId", "name", "CanView"
    
    Dim protectedRange As Object
    Set protectedRange = oWorksheet.GetProtectedRange("protectedRange") ' Retrieve the protected range
    
    Dim userInfo As Object
    Set userInfo = protectedRange.GetUser("userId") ' Get user information for "userId"
    
    Dim userName As String
    userName = userInfo.GetName() ' Get the user's name
    
    ' Set the value of cell A3 to display the user's name
    oWorksheet.Range("A3").Value = "Name: " & userName
End Sub
```