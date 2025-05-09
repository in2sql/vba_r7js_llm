```plaintext
// This example changes the user protected range.
// Этот пример изменяет защищенный диапазон пользователя.
```

```vba
' VBA code to change the user protected range

Sub ChangeUserProtectedRange()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a protected range named "protectedRange" covering cells A1:B1
    oWorksheet.Range("A1:B1").Name = "protectedRange"
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Note: Excel VBA does not support adding or deleting specific users for protected ranges.
    ' Protection is applied to the sheet or workbook level.
End Sub
```

```javascript
// This example changes the user protected range.
// Этот пример изменяет защищенный диапазон пользователя.

var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" in cells A1:B1
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1");

// Retrieve the protected range object
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Add a user with ID "userId", name "name", and permission "CanView"
protectedRange.AddUser("userId", "name", "CanView");

// Delete the user with ID "userId"
protectedRange.DeleteUser("userId");
```