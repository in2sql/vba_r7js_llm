**Example demonstrating setting a cell value, adding a comment with a reply, and retrieving the reply author's name.**  
**Пример демонстрирует установку значения ячейки, добавление комментария с ответом и получение имени автора ответа.**

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Add a reply to the comment
oComment.AddReply("Reply 1", "John Smith", "uid-1");

// Get the reply
var oReply = oComment.GetReply();

// Set label in cell A3
oWorksheet.GetRange("A3").SetValue("Comment's reply author: ");

// Set the reply author's name in cell B3
oWorksheet.GetRange("B3").SetValue(oReply.GetAuthorName());
```

```vba
' VBA code equivalent to the OnlyOffice JavaScript example

Sub AddCommentAndReply()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    Dim oComment As Comment
    Set oComment = oWorksheet.Range("A1").AddComment("This is just a number.")
    
    ' Add a reply to the comment
    ' VBA does not support comment replies directly, so we append the reply to the comment text
    oComment.Text Text:=oComment.Text & vbCrLf & "Reply 1 by John Smith (uid-1)"
    
    ' Set label in cell A3
    oWorksheet.Range("A3").Value = "Comment's reply author: "
    
    ' Retrieve the reply author's name from the comment
    Dim commentText As String
    commentText = oComment.Text
    Dim replyAuthor As String
    ' Extract the author's name from the appended reply
    replyAuthor = Split(Split(commentText, " by ")(1), " (")(0)
    
    ' Set the reply author's name in cell B3
    oWorksheet.Range("B3").Value = replyAuthor
End Sub
```