### Description
**English:** This example sets the timestamp of the comment reply creation in UTC format.
  
**Russian:** Этот пример устанавливает временную метку создания ответа на комментарий в формате UTC.

```vba
' VBA code equivalent to the OnlyOffice JS code

Sub SetCommentReplyTimestamp()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1
    Dim oComment As Comment
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Add a reply to the comment
    oComment.Reply Text:="Reply 1", Author:="John Smith"
    
    ' Set the UTC timestamp to current time
    ' VBA does not have a direct method to set UTC time for comments
    ' This is a placeholder to illustrate the intent
    Dim currentTime As Date
    currentTime = Now ' Replace with UTC time retrieval if necessary
    oComment.Date = currentTime
    
    ' Set value in cell A3
    oWorksheet.Range("A3").Value = "Comment's reply timestamp UTC: "
    
    ' Set value in cell B3 to reply timestamp
    oWorksheet.Range("B3").Value = oComment.Date
End Sub
```

```javascript
// This example sets the timestamp of the comment reply creation in UTC format.

function SetCommentReplyTimestamp() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value "1" in cell A1
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get the range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to A1
    var oComment = oRange.AddComment("This is just a number.");
    
    // Add a reply to the comment
    oComment.AddReply("Reply 1", "John Smith", "uid-1");
    
    // Get the reply
    var oReply = oComment.GetReply();
    
    // Set the UTC timestamp to current time
    oReply.SetTimeUTC(Date.now());
    
    // Set value in cell A3
    oWorksheet.GetRange("A3").SetValue("Comment's reply timestamp UTC: ");
    
    // Set value in cell B3 to reply timestamp
    oWorksheet.GetRange("B3").SetValue(oReply.GetTimeUTC());
}
```