## Description / Описание

**English:** This code sets a value in cell A1, adds a comment and a reply, retrieves the timestamp of the reply in UTC format, and sets it in cell B3.

**Русский:** Этот код устанавливает значение в ячейке A1, добавляет комментарий и ответ, получает отметку времени ответа в формате UTC и устанавливает ее в ячейке B3.

```vba
' VBA Code Equivalent

Sub AddCommentWithReplyTimestamp()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As CommentThreaded
    Dim oReply As CommentThreaded
    Dim replyTimestampUTC As String
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a threaded comment to A1
    Set oComment = oRange.AddCommentThreaded("This is just a number.", "Author")
    
    ' Add a reply to the comment
    Set oReply = oComment.Replies.Add("Reply 1", "John Smith")
    
    ' Assuming the reply has a DateTime property in UTC
    replyTimestampUTC = Format(oReply.DateTimeUTC, "yyyy-mm-dd HH:MM:SS")
    
    ' Set description text in cell A3
    oWorksheet.Range("A3").Value = "Comment's reply timestamp UTC: "
    
    ' Set the timestamp in cell B3
    oWorksheet.Range("B3").Value = replyTimestampUTC
End Sub
```

```javascript
// OnlyOffice JS Code Equivalent

// This example shows how to get the timestamp of the comment reply creation in UTC format.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set value "1" in cell A1
var oRange = oWorksheet.GetRange("A1"); // Get the range A1
var oComment = oRange.AddComment("This is just a number."); // Add a comment to A1
oComment.AddReply("Reply 1", "John Smith", "uid-1"); // Add a reply to the comment
var oReply = oComment.GetReply(); // Get the reply
oWorksheet.GetRange("A3").SetValue("Comment's reply timestamp UTC: "); // Set description in A3
oWorksheet.GetRange("B3").SetValue(oReply.GetTimeUTC()); // Set the UTC timestamp in B3
```