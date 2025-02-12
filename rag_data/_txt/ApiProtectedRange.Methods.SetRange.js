**English:**  
This code modifies a user-protected range in an Excel sheet by adding a protected range and then setting a new range for protection.

**Russian:**  
Этот код изменяет защищенный диапазон пользователя в листе Excel, добавляя защищенный диапазон и затем устанавливая новый диапазон для защиты.

```vba
' VBA Code to modify a user-protected range in Excel

Sub ModifyProtectedRange()
    Dim ws As Worksheet
    Dim protectedRange As Range
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Unprotect the sheet to modify protection settings
    ws.Unprotect Password:="your_password"
    
    ' Define the initial protected range
    Set protectedRange = ws.Range("$A$1:$B$1")
    
    ' Lock the initial range
    protectedRange.Locked = True
    
    ' Protect the sheet with user interface only
    ws.Protect Password:="your_password", UserInterfaceOnly:=True
    
    ' Define the new range to be protected
    Set protectedRange = ws.Range("$A$2:$B$2")
    
    ' Lock the new range
    protectedRange.Locked = True
    
    ' Protect the sheet again to apply changes
    ws.Protect Password:="your_password", UserInterfaceOnly:=True
End Sub
```

```javascript
// JavaScript Code to modify a user-protected range in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" to the specified cells
oWorksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");

// Retrieve the protected range by name
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Set a new range for the protected range
protectedRange.SetRange("Sheet1!$A$2:$B$2");
```