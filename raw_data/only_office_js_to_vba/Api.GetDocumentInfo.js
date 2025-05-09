**Description / Описание**

This code retrieves the document information and sets the application name in cell A1.
Этот код получает информацию о документе и устанавливает название приложения в ячейку A1.

**VBA Code:**

```vba
' This VBA code retrieves document information and sets the application name in cell A1

Sub SetApplicationName()
    ' Get document information
    Dim docInfo As Object
    Set docInfo = Api.GetDocumentInfo()
    
    ' Get the active sheet and cell A1
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").Value = "This document has been created with: " & docInfo.Application
End Sub
```

**OnlyOffice JS Code:**

```javascript
// This example shows how to get the document info represented as an object and paste the application name into "A1" cell.
const oDocInfo = Api.GetDocumentInfo();
const oRange = Api.GetActiveSheet().GetRange('A1');
oRange.SetValue('This document has been created with: ' + oDocInfo.Application);
```