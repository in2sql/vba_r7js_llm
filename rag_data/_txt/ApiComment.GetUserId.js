# Description
This code sets the value of cell A1 to "1", adds a comment to A1, sets cell A3 with a label, and retrieves the user ID of the comment's author.
Этот код устанавливает значение ячейки A1 равным "1", добавляет комментарий к A1, устанавливает метку в ячейке A3 и получает идентификатор пользователя автора комментария.

## VBA Code
```vba
' This VBA code sets a value in A1, adds a comment, and retrieves the comment author's name.
' Этот VBA-код устанавливает значение в A1, добавляет комментарий и получает имя автора комментария.

Sub AddCommentAndGetAuthor()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value in A1
    ws.Range("A1").Value = "1"
    
    ' Add a comment to A1
    Set rng = ws.Range("A1")
    
    ' Remove existing comment if any
    If Not rng.Comment Is Nothing Then
        rng.Comment.Delete
    End If
    
    ' Add new comment
    Set cmt = rng.AddComment("This is just a number.")
    
    ' Set the label in A3
    ws.Range("A3").Value = "Comment's user name:"
    
    ' Set the author's name in B3
    ws.Range("B3").Value = cmt.Author
End Sub
```

## JavaScript Code
```javascript
// This example shows how to set a value, add a comment, and get the user ID of the comment author.
// Этот пример показывает, как установить значение, добавить комментарий и получить идентификатор пользователя автора комментария.

function addCommentAndGetUserId() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value in A1
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get the range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to A1
    var oComment = oRange.AddComment("This is just a number.");
    
    // Set the label in A3
    oWorksheet.GetRange("A3").SetValue("Comment's user Id: ");
    
    // Set the user's ID in B3
    oWorksheet.GetRange("B3").SetValue(oComment.GetUserId());
}
```