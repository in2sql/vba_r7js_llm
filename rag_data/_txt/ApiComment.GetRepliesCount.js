**Description / Описание**

This code demonstrates how to add a comment to a cell, add a reply to that comment, and then retrieve and display the count of replies.
Этот код демонстрирует, как добавить комментарий к ячейке, добавить ответ к этому комментарию и затем получить и отобразить количество ответов.

```vba
' VBA Code to add a comment, add a reply, and get the count of replies

Sub AddCommentAndReplies()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    Dim replyCount As Integer
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Add a reply to the comment
    oComment.Replies.Add "Reply 1 by John Smith"
    
    ' Set value in cell A3
    oWorksheet.Range("A3").Value = "Comment replies count: "
    
    ' Get the count of replies
    replyCount = oComment.Replies.Count
    
    ' Set the count in cell B3
    oWorksheet.Range("B3").Value = replyCount
End Sub
```

```javascript
// JavaScript Code to add a comment, add a reply, and get the count of replies using OnlyOffice API

function addCommentAndReplies() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value in cell A1
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get the range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to A1
    var oComment = oRange.AddComment("This is just a number.");
    
    // Add a reply to the comment
    oComment.AddReply("Reply 1", "John Smith", "uid-1");
    
    // Set value in cell A3
    oWorksheet.GetRange("A3").SetValue("Comment replies count: ");
    
    // Get the count of replies and set it in cell B3
    oWorksheet.GetRange("B3").SetValue(oComment.GetRepliesCount());
}
```