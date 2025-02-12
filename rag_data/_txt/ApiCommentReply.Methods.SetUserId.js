# Description / Описание

This code sets the user ID for the comment reply author in an OnlyOffice spreadsheet.

Этот код устанавливает идентификатор пользователя для автора ответа на комментарий в электронной таблице OnlyOffice.

## VBA Code

```vba
' VBA code to set the user ID for the comment reply author

Sub SetCommentReplyAuthor()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    ' Note: VBA may not support replies and setting user IDs directly
    
    Set oWorksheet = ThisWorkbook.ActiveSheet
    oWorksheet.Range("A1").Value = "1" ' Set value "1" in cell A1
    
    Set oRange = oWorksheet.Range("A1") ' Get range A1
    Set oComment = oRange.AddComment("This is just a number.") ' Add comment
    
    ' VBA does not support adding replies with user IDs directly
    ' Additional implementation required for custom properties
    
    oWorksheet.Range("A3").Value = "Comment's reply user Id: "
    ' The following line is a placeholder as VBA does not support GetUserId
    ' oWorksheet.Range("B3").Value = oReply.UserId
End Sub
```

## JavaScript Code

```javascript
// This example sets the user ID to the comment reply author.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set value "1" in cell A1

var oRange = oWorksheet.GetRange("A1"); // Get the range A1
var oComment = oRange.AddComment("This is just a number."); // Add a comment to A1

oComment.AddReply("Reply 1", "John Smith", "uid-1"); // Add a reply to the comment

var oReply = oComment.GetReply(); // Get the reply
oReply.SetUserId("uid-2"); // Set the user ID of the reply

oWorksheet.GetRange("A3").SetValue("Comment's reply user Id: "); // Set label in A3
oWorksheet.GetRange("B3").SetValue(oReply.GetUserId()); // Set the user ID in B3
```