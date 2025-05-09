**Description:**
English: This code changes the title of a user-protected range in the active worksheet.
  
Russian: Этот код изменяет заголовок пользовательского защищенного диапазона на активном листе.

```vba
' VBA Code: Change the title of a protected range in the active worksheet

Sub ChangeProtectedRangeTitle()
    Dim ws As Worksheet
    Dim pr As AllowEditRange
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Add a protected range named "protectedRange" covering cells A1:B1
    Set pr = ws.Protection.AllowEditRanges.Add(Title:="protectedRange", Range:=ws.Range("A1:B1"))
    
    ' Change the title of the protected range to "protectedRangeNew"
    pr.Title = "protectedRangeNew"
End Sub
```

```javascript
// JavaScript Code: Change the title of a protected range in the active worksheet

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a protected range named "protectedRange" covering cells A1:B1 on Sheet1
oWorksheet.AddProtectedRange("protectedRange", "Sheet1!$A$1:$B$1");

// Retrieve the protected range
var protectedRange = oWorksheet.GetProtectedRange("protectedRange");

// Set the new title for the protected range
protectedRange.SetTitle("protectedRangeNew");
```