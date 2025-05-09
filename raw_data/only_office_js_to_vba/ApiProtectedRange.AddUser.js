**Description / Описание**

This code changes the user protected range in the active worksheet.

Этот код изменяет защищенный диапазон пользователя в активном рабочем листе.

```javascript
// JavaScript code using OnlyOffice API
// This example changes the user protected range.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.AddProtectedRange("protectedRange", "$A$1:$B$1");
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");
protectedRange.AddUser("userId", "name", "CanView"); 
```

```vba
' VBA code equivalent
' This example changes the user protected range.

Sub ChangeUserProtectedRange()
    Dim ws As Worksheet
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Unprotect the worksheet if it's already protected
    ws.Unprotect Password:="password"
    
    ' Define the range to protect
    With ws.Range("A1:B1")
        .Locked = True ' Lock the range
    End With
    
    ' Protect the worksheet with a password
    ' VBA does not support adding specific users to protected ranges directly
    ws.Protect Password:="password", UserInterfaceOnly:=True
End Sub
```