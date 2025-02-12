**Description / Описание**

English: This code adds a protected range to the active sheet, assigns a user with view permissions, retrieves the user ID, and writes it into cell A3.

Russian: Этот код добавляет защищенный диапазон на активный лист, назначает пользователя с правами просмотра, получает идентификатор пользователя и записывает его в ячейку A3.

```vba
' This VBA code adds protection to a specified range, allows a user to view it, retrieves the user ID, and writes it to cell A3.

Sub ProtectRangeAndAddUser()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define the range to protect
    With ws
        .Unprotect ' Unprotect the sheet to modify protection settings
        .Range("A1:B1").Locked = True
        .Protect Password:="password", UserInterfaceOnly:=True
    End With
    
    ' Note: Excel VBA does not support per-user permissions on protected ranges.
    ' This functionality would require custom implementation or external tools.
    
    ' For demonstration, writing a mock user ID to cell A3
    Dim userId As String
    userId = "userId123" ' Replace with actual user ID retrieval method
    ws.Range("A3").Value = "Id: " & userId
End Sub
```

```javascript
// This OnlyOffice JS code adds a protected range to the active sheet, assigns a user with view permissions, retrieves the user ID, and writes it into cell A3.

function protectRangeAndAddUser() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Add a protected range and assign a user with 'CanView' permission
    oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");
    
    // Retrieve the protected range
    var protectedRange = oWorksheet.GetProtectedRange("protectedRange");
    
    // Get user information
    var userInfo = protectedRange.GetUser("userId");
    
    // Get user ID
    var userId = userInfo.GetId();
    
    // Write the user ID to cell A3
    oWorksheet.GetRange("A3").SetValue("Id: " + userId);
}
```