**Description:**
*English:* This code demonstrates how to add a comment to the document, retrieve it by ID, and set specific cell values with the comment text and author.
*Russian:* Этот код демонстрирует, как добавить комментарий в документ, получить его по ID и установить значения определенных ячеек с текстом комментария и автором.

```javascript
// This code demonstrates adding a comment, retrieving it by ID, and setting cell values with comment text and author.

var oComment = Api.AddComment("Comment", "Bob");
var sId = oComment.GetId();
oComment = Api.GetCommentById(sId);
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("Comment Text: ", oComment.GetText());
oWorksheet.GetRange("B1").SetValue("Comment Author: ", oComment.GetAuthorName());
```

```vba
' This VBA code demonstrates adding a comment, retrieving it by ID, and setting cell values with comment text and author.

Sub AddAndRetrieveComment()
    Dim oComment As Object
    Dim sId As String
    Dim oWorksheet As Object

    ' Add a comment with text "Comment" and author "Bob"
    Set oComment = Api.AddComment("Comment", "Bob")
    
    ' Get the ID of the comment
    sId = oComment.GetId()
    
    ' Retrieve the comment by its ID
    Set oComment = Api.GetCommentById(sId)
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Set cell A1 with the comment text
    oWorksheet.GetRange("A1").SetValue "Comment Text: ", oComment.GetText()
    
    ' Set cell B1 with the comment author name
    oWorksheet.GetRange("B1").SetValue "Comment Author: ", oComment.GetAuthorName()
End Sub
```