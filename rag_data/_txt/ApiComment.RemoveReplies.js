**Description:**
English: This example removes the specified comment replies.
Russian: Этот пример удаляет указанные ответы на комментарии.

```javascript
// This example removes the specified comment replies
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set value of cell A1 to "1"
var oRange = oWorksheet.GetRange("A1"); // Get the range A1
var oComment = oRange.AddComment("This is just a number."); // Add a comment to A1
oComment.AddReply("Reply 1", "John Smith", "uid-1"); // Add first reply to the comment
oComment.AddReply("Reply 2", "John Smith", "uid-1"); // Add second reply to the comment
oComment.RemoveReplies(0, 1, false); // Remove the first reply
oWorksheet.GetRange("A3").SetValue("Comment replies count: "); // Set cell A3 text
oWorksheet.GetRange("B3").SetValue(oComment.GetRepliesCount()); // Set cell B3 to the number of comment replies
```

```vba
' This example removes the specified comment replies
' Этот пример удаляет указанные ответы на комментарии

Sub RemoveCommentReplies()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet ' Get the active worksheet
    
    oWorksheet.Range("A1").Value = "1" ' Set value of cell A1 to "1"
    
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1") ' Get the range A1
    
    Dim oComment As ThreadedComment
    Set oComment = oRange.AddCommentThreaded("This is just a number.") ' Add a comment to A1
    
    oComment.Replies.Add "Reply 1", "John Smith", "uid-1" ' Add first reply to the comment
    oComment.Replies.Add "Reply 2", "John Smith", "uid-1" ' Add second reply to the comment
    
    oComment.Replies.Remove 0, 1, False ' Remove the first reply
    
    oWorksheet.Range("A3").Value = "Comment replies count: " ' Set cell A3 text
    oWorksheet.Range("B3").Value = oComment.Replies.Count ' Set cell B3 to the number of comment replies
End Sub
```