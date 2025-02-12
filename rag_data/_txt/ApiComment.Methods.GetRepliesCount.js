# Get and Count Comment Replies / Получение и подсчет ответов на комментарий

This script sets a value in cell A1, adds a comment to it, adds a reply to the comment, and then displays the count of replies in cells A3 and B3.

Этот скрипт устанавливает значение в ячейке A1, добавляет к ней комментарий, добавляет ответ к комментарию, а затем отображает количество ответов в ячейках A3 и B3.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range object for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Add a reply to the comment
oComment.AddReply("Reply 1", "John Smith", "uid-1");

// Set the value of cell A3 to the description
oWorksheet.GetRange("A3").SetValue("Comment replies count: ");

// Get the count of comment replies and set it in cell B3
oWorksheet.GetRange("B3").SetValue(oComment.GetRepliesCount());
```

```vba
' VBA code equivalent

Sub GetCommentRepliesCount()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range object for cell A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to cell A1
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Simulate adding a reply by appending text to the comment
    oComment.Text Text:=oComment.Text & vbCrLf & "Reply 1 by John Smith"
    
    ' Set the value of cell A3 to the description
    oWorksheet.Range("A3").Value = "Comment replies count: "
    
    ' Simulate getting the replies count, assuming 1 reply
    oWorksheet.Range("B3").Value = 1
End Sub
```