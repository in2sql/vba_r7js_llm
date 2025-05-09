**Description: Sets the user ID to the comment author.
Описание: Устанавливает идентификатор пользователя для автора комментария.**

```vba
' VBA Code: Sets the user ID to the comment author

Sub SetCommentUserId()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    ' Get the active sheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1 with the text and author
    Set oComment = oRange.AddComment("This is just a number.")
    oComment.Author = "John Smith"
    
    ' Set the value of cell A3
    oWorksheet.Range("A3").Value = "Comment's user Id: "
    
    ' Store the User ID in the comment's text
    oComment.Shape.TextFrame.Characters.Text = oComment.Shape.TextFrame.Characters.Text & vbCrLf & "UserId: uid-2"
    
    ' Set the value of cell B3 to the User ID
    oWorksheet.Range("B3").Value = "uid-2"
End Sub
```

```javascript
// OnlyOffice JS Code: Sets the user ID to the comment author

// This example sets the user ID to the comment author.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
var oComment = oRange.AddComment("This is just a number.", "John Smith");
oWorksheet.GetRange("A3").SetValue("Comment's user Id: ");
oComment.SetUserId("uid-2");
oWorksheet.GetRange("B3").SetValue(oComment.GetUserId()); 
```