## Example: Retrieve Comment Reply Author's Name
Этот пример демонстрирует, как получить имя автора ответа на комментарий.

### VBA Code
```vba
' VBA code to retrieve the comment reply author's name
Sub GetCommentReplyAuthor()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to cell A1
    Dim oComment As Comment
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Add a reply to the comment
    oComment.Replies.Add "Reply 1", "John Smith", "uid-1"
    
    ' Get the first reply
    Dim oReply As CommentReply
    Set oReply = oComment.Replies(1)
    
    ' Set the value of cell A3
    oWorksheet.Range("A3").Value = "Comment's reply author: "
    
    ' Set the value of cell B3 to the author's name
    oWorksheet.Range("B3").Value = oReply.Author
End Sub
```

### JavaScript Code
```javascript
// JavaScript code to retrieve the comment reply author's name
function getCommentReplyAuthor() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set the value of cell A1 to "1"
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get the range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to cell A1
    var oComment = oRange.AddComment("This is just a number.");
    
    // Add a reply to the comment
    oComment.AddReply("Reply 1", "John Smith", "uid-1");
    
    // Get the first reply
    var oReply = oComment.GetReply();
    
    // Set the value of cell A3
    oWorksheet.GetRange("A3").SetValue("Comment's reply author: ");
    
    // Set the value of cell B3 to the author's name
    oWorksheet.GetRange("B3").SetValue(oReply.GetAuthorName());
}
```