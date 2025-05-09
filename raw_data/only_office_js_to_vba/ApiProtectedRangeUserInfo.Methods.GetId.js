### Description / Описание
**English:** This script modifies a protected range in the active worksheet by adding a user with specific permissions and displays the user's ID in cell A3.

**Russian:** Этот скрипт изменяет защищенный диапазон в активном листе, добавляя пользователя с определенными разрешениями и отображает ID пользователя в ячейке A3.

```javascript
// JavaScript OnlyOffice API code to modify a protected range and add a user

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" covering cells A1 to B1 and add a user with view permissions
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1").AddUser("userId", "name", "CanView");

// Retrieve the protected range
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Get information about the user with ID "userId"
var userInfo = protectedRange.GetUser("userId");

// Get the user's ID
var userId = userInfo.GetId();

// Set the value of cell A3 to display the user's ID
oWorksheet.GetRange("A3").SetValue("Id: " + userId);
```

```vba
' VBA code to protect a range, add a user with specific permissions, and display the user ID

Sub ModifyProtectedRange()
    Dim oWorksheet As Worksheet
    Dim userId As String
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Protect the range A1:B1
    With oWorksheet.Range("A1:B1").Protection
        .AllowSelectingLockedCells = False
        .AllowFormattingCells = False
        .AllowUsingPivotTables = False
        .Protect Password:="yourPassword", UserInterfaceOnly:=True
    End With
    
    ' Note: Excel VBA does not support adding specific users to protected ranges.
    ' Permissions are generally managed via workbook protection with a password.
    
    ' Set the user ID (this is a placeholder as VBA cannot retrieve user IDs in this context)
    userId = "userId"
    
    ' Display the user ID in cell A3
    oWorksheet.Range("A3").Value = "Id: " & userId
End Sub
```