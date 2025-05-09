## Description / Описание

**English:** This code adds a reply to a comment in an OnlyOffice spreadsheet.

**Russian:** Этот код добавляет ответ к комментарию в электронной таблице OnlyOffice.

```vba
' VBA code to add a comment and a reply to a cell, then display the reply text

Sub AddCommentWithReply()
    Dim ws As Worksheet
    Dim rng As Range
    Dim commentObj As Comment
    Dim replyText As String
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Set value in cell A1
    ws.Range("A1").Value = "1"
    
    ' Add comment to cell A1
    Set rng = ws.Range("A1")
    Set commentObj = rng.AddComment("This is just a number.")
    
    ' Add a reply to the comment by appending text
    commentObj.Text Text:=commentObj.Text & vbCrLf & "Reply 1 - John Smith (uid-1)"
    
    ' Retrieve the reply text
    replyText = "Reply 1 - John Smith (uid-1)"
    
    ' Set value in cells A3 and B3
    ws.Range("A3").Value = "Comment's reply text: "
    ws.Range("B3").Value = replyText
End Sub
```

```javascript
// JavaScript code to add a comment and a reply to a cell, then display the reply text

function AddCommentWithReply() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value in cell A1
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Add comment to cell A1
    var oRange = oWorksheet.GetRange("A1");
    var oComment = oRange.AddComment("This is just a number.");
    
    // Add a reply to the comment
    oComment.AddReply("Reply 1", "John Smith", "uid-1");
    
    // Retrieve the reply text
    var oReply = oComment.GetReply();
    
    // Set value in cells A3 and B3
    oWorksheet.GetRange("A3").SetValue("Comment's reply text: ");
    oWorksheet.GetRange("B3").SetValue(oReply.GetText());
}
```