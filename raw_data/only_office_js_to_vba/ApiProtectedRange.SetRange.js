# Description / Описание

**English**: This code modifies a user-protected range by adding it to the worksheet and then changing its range from A1:B1 to A2:B2.

**Russian**: Этот код изменяет пользовательский защищенный диапазон, добавляя его в лист и затем изменяя его диапазон с A1:B1 на A2:B2.

```vba
' VBA Code to add and modify a protected range
Sub ModifyProtectedRange()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim protectedName As String
    
    ' Set the worksheet to the active sheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    protectedName = "protectedRange"
    
    ' Add a named range for protection
    Set oRange = oWorksheet.Range("A1:B1")
    ThisWorkbook.Names.Add Name:=protectedName, RefersTo:=oRange
    
    ' Protect the worksheet, allowing only the named range to be edited
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True
    
    ' Unprotect the worksheet to modify the protected range
    oWorksheet.Unprotect Password:="password"
    
    ' Update the named range to new cells
    ThisWorkbook.Names(protectedName).RefersTo = "=" & oWorksheet.Name & "!$A$2:$B$2"
    
    ' Protect the worksheet again
    oWorksheet.Protect Password:="password", UserInterfaceOnly:=True
End Sub
```

```javascript
// OnlyOffice JS Code to add and modify a protected range
function modifyProtectedRange() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Add a protected range named "protectedRange" to A1:B1
    oWorksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");
    
    // Get the protected range object
    var protectedRange = oWorksheet.GetProtectedRange("protectedRange");
    
    // Change the range to A2:B2
    protectedRange.SetRange("Sheet1!$A$2:$B$2");
}
```