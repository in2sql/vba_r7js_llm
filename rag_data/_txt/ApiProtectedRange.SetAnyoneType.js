# This example changes the user protected range.
# Этот пример изменяет защищенный диапазон пользователя.

```vba
' VBA code to change the user protected range
Sub ChangeProtectedRange()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet ' Get the active sheet
    
    ' Unprotect the sheet to modify protection settings
    oWorksheet.Unprotect Password:="password"
    
    ' Add protected range by locking cells A1:B1
    oWorksheet.Range("A1:B1").Locked = True
    
    ' Protect the sheet, allowing users to select unlocked cells only
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True, AllowSelectingLockedCells:=False
End Sub
```

```javascript
// This example changes the user protected range.
var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
oWorksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1"); // Add a protected range
var protectedRange = oWorksheet.GetProtectedRange("protectedRange"); // Get the protected range
protectedRange.SetAnyoneType("CanView"); // Set permission to view only
```