```javascript
// This script adds a protected range to the active sheet and sets its permission to view only.
// Этот скрипт добавляет защищенный диапазон на активный лист и устанавливает его разрешения только для просмотра.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" covering cells A1 to B1 on Sheet1
oWorksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");

// Retrieve the protected range just added
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Set the permission type of the protected range to "CanView"
protectedRange.SetAnyoneType("CanView"); 
```

```vba
' This macro adds a protected range to "Sheet1" and sets its permission to view only.
' Этот макрос добавляет защищенный диапазон на "Лист1" и устанавливает его разрешения только для просмотра.

Sub AddProtectedRange()
    Dim ws As Worksheet
    Dim protectedRange As Range
    Dim protection As Protection
    
    ' Set the worksheet to Sheet1
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Define the range A1:B1
    Set protectedRange = ws.Range("A1:B1")
    
    ' Protect the worksheet if it's not already protected
    If Not ws.ProtectContents Then
        ws.Protect Password:="", UserInterfaceOnly:=True
    End If
    
    ' Lock the specified range
    protectedRange.Locked = True
    
    ' Optionally, you can add further permission settings here
    ' Note: VBA does not have a direct equivalent of "CanView", 
    ' so worksheet protection is used to restrict editing.
End Sub
```