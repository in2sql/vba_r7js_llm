**Description / Описание**

This code manipulates comments in an Excel sheet by adding comments, retrieving them, and setting cell values based on the comments.

Этот код управляет комментариями в листе Excel, добавляя комментарии, извлекая их и устанавливая значения ячеек на основе комментариев.

```vba
' VBA Code to manage comments in Excel

Sub ManageComments()
    Dim oWorksheet As Worksheet
    Dim arrComments As Comments
    Dim commentText As String
    Dim commentAuthor As String
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Add a global comment with text "Comment 1" and author "John Smith"
    oWorksheet.Comments.Add Range:=oWorksheet.Range("A1"), Text:="Comment 1"
    oWorksheet.Range("A1").Comment.Author = "John Smith"
    
    ' Add a comment to cell A4 with text "Comment 2" and author "Mark Potato"
    oWorksheet.Range("A4").AddComment Text:="Comment 2"
    oWorksheet.Range("A4").Comment.Author = "Mark Potato"
    
    ' Retrieve all comments from the worksheet
    Set arrComments = oWorksheet.Comments
    
    ' Ensure there are at least two comments
    If arrComments.Count >= 2 Then
        ' Get the text of the second comment
        commentText = arrComments(2).Text
        
        ' Get the author of the second comment
        commentAuthor = arrComments(2).Author
        
        ' Set cell A1 with the comment text
        oWorksheet.Range("A1").Value = "Comment text: " & commentText
        
        ' Set cell A2 with the comment author
        oWorksheet.Range("A2").Value = "Comment author: " & commentAuthor
    Else
        MsgBox "Not enough comments to retrieve information.", vbExclamation
    End If
End Sub
```

```javascript
// JavaScript Code to manage comments in OnlyOffice

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Add a global comment with text "Comment 1" and author "John Smith"
Api.AddComment("Comment 1", "John Smith");

// Add a comment to cell A4 with text "Comment 2" and author "Mark Potato"
oWorksheet.GetRange("A4").AddComment("Comment 2", "Mark Potato");

// Retrieve all comments from the worksheet
var arrComments = Api.GetAllComments();

// Ensure there are at least two comments
if (arrComments.length >= 2) {
    // Get the text of the second comment
    var commentText = arrComments[1].GetText();
    
    // Get the author of the second comment
    var commentAuthor = arrComments[1].GetAuthorName();
    
    // Set cell A1 with the comment text
    oWorksheet.GetRange("A1").SetValue("Comment text: " + commentText);
    
    // Set cell A2 with the comment author
    oWorksheet.GetRange("A2").SetValue("Comment author: " + commentAuthor);
} else {
    // Alert if there are not enough comments
    Api.ShowMessage("Not enough comments to retrieve information.", "Warning");
}
```