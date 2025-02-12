---

**Description:**

- **English:** This example retrieves the document information and writes the application name into cell A1.
- **Russian:** Этот пример получает информацию о документе и записывает имя приложения в ячейку A1.

```vba
' VBA code
' This example retrieves the document information and writes the application name into cell A1.

Sub SetApplicationName()
    ' Get document information
    Dim docInfo As DocumentInfo
    Set docInfo = Api.GetDocumentInfo()
    
    ' Get range A1
    Dim rng As Range
    Set rng = ActiveSheet.Range("A1")
    
    ' Set value to cell A1
    rng.Value = "This document has been created with: " & docInfo.Application
End Sub
```

```javascript
// JavaScript code
// This example retrieves the document information and writes the application name into cell A1.

const docInfo = Api.GetDocumentInfo(); // Retrieve document information
const range = Api.GetActiveSheet().GetRange('A1'); // Get range A1
range.SetValue('This document has been created with: ' + docInfo.Application); // Set value
```