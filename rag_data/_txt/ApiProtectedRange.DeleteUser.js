### Description / Описание

**English:**  
This code modifies the user permissions for a protected range in the active worksheet. It adds a protected range from cell A1 to B1, assigns a user with view permissions, and then removes the user from the protected range.

**Russian:**  
Этот код изменяет разрешения пользователя для защищенного диапазона на активном листе. Он добавляет защищенный диапазон от ячейки A1 до B1, назначает пользователя с правами просмотра, а затем удаляет пользователя из защищенного диапазона.

```vba
' VBA code to modify user permissions for a protected range

Sub ModifyProtectedRange()
    Dim oWorksheet As Worksheet
    Dim protectedRange As Range
    Dim userId As String
    Dim userName As String
    
    ' Set the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define the protected range
    Set protectedRange = oWorksheet.Range("A1:B1")
    
    ' Protect the range
    protectedRange.Locked = True
    oWorksheet.Protect Password:="yourPassword", UserInterfaceOnly:=True
    
    ' Add user permissions (VBA does not support user-specific permissions directly)
    ' This is a placeholder as Excel VBA has limited support for user-specific range protection
    MsgBox "Excel VBA does not support adding or deleting users for protected ranges directly."
    
    ' To remove protection
    ' oWorksheet.Unprotect Password:="yourPassword"
End Sub
```

```javascript
// JavaScript code to modify user permissions for a protected range using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" covering cells A1 to B1
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1");

// Retrieve the protected range object
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Add a user with view permissions to the protected range
protectedRange.AddUser("userId", "name", "CanView");

// Remove the user from the protected range
protectedRange.DeleteUser("userId");
```