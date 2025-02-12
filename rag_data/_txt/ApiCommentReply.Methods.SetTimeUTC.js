# Code Description / Описание кода

This code sets the value of cell A1 to "1", adds a comment to cell A1 stating "This is just a number.", adds a reply to the comment from "John Smith", sets the reply's timestamp in UTC, and then writes the timestamp to cells A3 and B3.

Этот код устанавливает значение ячейки A1 на "1", добавляет комментарий к ячейке A1 со ссылкой "This is just a number.", добавляет ответ на комментарий от "John Smith", устанавливает временную метку ответа в UTC, а затем записывает временную метку в ячейки A3 и B3.

```javascript
// JavaScript Code using OnlyOffice API

// This example sets the timestamp of the comment reply creation in UTC format.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
var oComment = oRange.AddComment("This is just a number.");
oComment.AddReply("Reply 1", "John Smith", "uid-1");
var oReply = oComment.GetReply();
oReply.SetTimeUTC(Date.now());
oWorksheet.GetRange("A3").SetValue("Comment's reply timestamp UTC: ");
oWorksheet.GetRange("B3").SetValue(oReply.GetTimeUTC());
```

```vba
' VBA Code equivalent using Excel VBA

' This example sets the timestamp of the comment reply creation in UTC format
Sub SetCommentReplyTimestampUTC()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    Dim oReply As Comment
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to cell A1
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Add a reply to the comment
    oComment.Replies.Add "Reply 1", "John Smith", "uid-1"
    
    ' Get the reply
    Set oReply = oComment.Replies.Item("Reply 1")
    
    ' Set the reply's timestamp to current UTC time
    ' Note: VBA does not directly support UTC, so you may need to adjust accordingly
    oReply.Date = Now ' This sets the server's local time
    
    ' Set the description in cell A3
    oWorksheet.Range("A3").Value = "Comment's reply timestamp UTC: "
    
    ' Set the timestamp value in cell B3
    oWorksheet.Range("B3").Value = oReply.Date
End Sub
```