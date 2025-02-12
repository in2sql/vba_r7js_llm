**Description / Описание:**
This example changes the user protected range.  
Этот пример изменяет защищенный диапазон пользователя.

```vba
' VBA code to add a protected range and assign user permissions
Sub AddProtectedRange()
    Dim oWorksheet As Worksheet
    Dim protectedRange As Range
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the protected range
    Set protectedRange = oWorksheet.Range("A1:B1")
    
    ' Lock the range
    protectedRange.Locked = True
    
    ' Protect the worksheet
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Note: VBA does not support adding specific users to protected ranges directly.
    ' User permissions need to be managed through the overall sheet protection.
End Sub
```

```javascript
// This example changes the user protected range.
// Этот пример изменяет защищенный диапазон пользователя.
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" covering cells A1 to B1
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1");

// Retrieve the protected range object
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Add a user with specific permissions to the protected range
protectedRange.AddUser("userId", "name", "CanView");
```