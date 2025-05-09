**Description: This code adds a reply to a comment in a worksheet.**

**Описание: Этот код добавляет ответ к комментарию на листе.**

```javascript
// This example adds a reply to a comment.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
var oComment = oRange.AddComment("This is just a number.");
oComment.AddReply("Reply 1", "John Smith", "uid-1");
var oReply = oComment.GetReply();
oWorksheet.GetRange("A3").SetValue("Comment's reply text: ");
oWorksheet.GetRange("B3").SetValue(oReply.GetText());
```

```vba
' This example adds a reply to a comment.
Sub AddCommentReply()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As CommentThreaded
    Dim oReply As CommentThreaded
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a threaded comment
    Set oComment = oRange.AddCommentThreaded("This is just a number.")
    
    ' Add a reply to the comment
    oComment.Replies.Add "Reply 1", "John Smith", "uid-1"
    
    ' Get the reply
    Set oReply = oComment.Replies.Item(1)
    
    ' Write reply text to cells
    oWorksheet.Range("A3").Value = "Comment's reply text: "
    oWorksheet.Range("B3").Value = oReply.Text
End Sub
```