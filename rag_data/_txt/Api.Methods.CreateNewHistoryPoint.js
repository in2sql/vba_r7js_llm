## Description / Описание

**English:** This example creates a new history point.

**Russian:** Этот пример создает новую точку истории.

```vba
' VBA Code
Sub CreateNewHistoryPoint()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value in A1
    oWorksheet.Range("A1").Value = "This is just a sample text."
    
    ' Create a new history point (Custom implementation required)
    ' VBA does not have a native method for history points like OnlyOffice
    ' This could be implemented using a versioning system or by saving the workbook state
    
    ' Set value in A3
    oWorksheet.Range("A3").Value = "New history point was just created."
End Sub
```

```javascript
// JavaScript Code
// This example creates a new history point.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("This is just a sample text.");
Api.CreateNewHistoryPoint();
oWorksheet.GetRange("A3").SetValue("New history point was just created.");
```