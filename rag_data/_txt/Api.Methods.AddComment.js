**Description / Описание**
This code adds comments to the document, retrieves them, and sets their text and author in specific cells of the active worksheet.
Этот код добавляет комментарии в документ, извлекает их и устанавливает их текст и автора в определенные ячейки активного рабочего листа.

```javascript
// This example adds comments to the document.
// Добавляет комментарии в документ.
Api.AddComment("Comment 1", "Bob"); // Adds a comment with author "Bob"
Api.AddComment("Comment 2"); // Adds a comment without specifying an author
var arrComments = Api.GetComments(); // Retrieves all comments
var oWorksheet = Api.GetActiveSheet(); // Gets the active worksheet
// Sets the text of the first comment in cell A1
oWorksheet.GetRange("A1").SetValue("Comment Text: ", arrComments[0].GetText());
// Sets the author of the first comment in cell B1
oWorksheet.GetRange("B1").SetValue("Comment Author: ", arrComments[0].GetAuthorName());
```

```vba
' This example adds comments to the document,
' retrieves them, and sets their text and author in specific cells of the active worksheet.
' Этот пример добавляет комментарии в документ, извлекает их и устанавливает их текст и автора в определенные ячейки активного рабочего листа.

Sub AddAndRetrieveComments()
    ' Adds a comment with author "Bob"
    Api.AddComment "Comment 1", "Bob"
    ' Adds a comment without specifying an author
    Api.AddComment "Comment 2"
    
    ' Retrieves all comments
    Dim arrComments As Variant
    Set arrComments = Api.GetComments()
    
    ' Gets the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Sets the text of the first comment in cell A1
    oWorksheet.GetRange("A1").SetValue "Comment Text: ", arrComments(0).GetText()
    
    ' Sets the author of the first comment in cell B1
    oWorksheet.GetRange("B1").SetValue "Comment Author: ", arrComments(0).GetAuthorName()
End Sub
```