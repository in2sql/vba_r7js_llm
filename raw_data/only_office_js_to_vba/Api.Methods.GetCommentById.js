# Description / Описание

**English:**  
This code adds a comment to the current document, retrieves it by its ID, and sets the comment's text and author into specific cells of the active worksheet.

**Russian:**  
Этот код добавляет комментарий к текущему документу, получает его по ID и записывает текст комментария и имя автора в определенные ячейки активного листа.

```vba
' VBA Code
' This code adds a comment to the current document, retrieves it by its ID, and sets
' the comment's text and author into specific cells of the active worksheet.

Sub AddAndRetrieveComment()
    Dim oComment As Object
    Dim sId As String
    Dim oWorksheet As Worksheet
    
    ' Add a comment with text "Comment" and author "Bob"
    Set oComment = Api.AddComment("Comment", "Bob")
    
    ' Get the ID of the comment
    sId = oComment.GetId()
    
    ' Retrieve the comment by its ID
    Set oComment = Api.GetCommentById(sId)
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Set the comment's text in cell A1
    oWorksheet.Range("A1").Value = "Comment Text: " & oComment.GetText()
    
    ' Set the comment's author in cell B1
    oWorksheet.Range("B1").Value = "Comment Author: " & oComment.GetAuthorName()
End Sub
```

```javascript
// OnlyOffice JS Code
// This code adds a comment to the current document, retrieves it by its ID, and sets
// the comment's text and author into specific cells of the active worksheet.

function AddAndRetrieveComment() {
    // Add a comment with text "Comment" and author "Bob"
    var oComment = Api.AddComment("Comment", "Bob");
    
    // Get the ID of the comment
    var sId = oComment.GetId();
    
    // Retrieve the comment by its ID
    oComment = Api.GetCommentById(sId);
    
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set the comment's text in cell A1
    oWorksheet.GetRange("A1").SetValue("Comment Text: " + oComment.GetText());
    
    // Set the comment's author in cell B1
    oWorksheet.GetRange("B1").SetValue("Comment Author: " + oComment.GetAuthorName());
}
```