# Description / Описание

**English:** This example sets the comment author's name.

**Russian:** Этот пример устанавливает имя автора комментария.

```vba
' Excel VBA code to set the comment author's name

Sub SetCommentAuthor()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1 with text
    Set oComment = oRange.AddComment("This is just a number.")
    oComment.Author = "John Smith"
    
    ' Set value "Comment's author:" in cell A3
    oWorksheet.Range("A3").Value = "Comment's author: "
    
    ' Change the author name of the comment
    oComment.Author = "Mark Potato"
    
    ' Set the value of B3 to the comment's author name
    oWorksheet.Range("B3").Value = oComment.Author
End Sub
```

```javascript
// This example sets the comment author's name.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1"); // Set value "1" in cell A1
var oRange = oWorksheet.GetRange("A1"); // Get the range A1
var oComment = oRange.AddComment("This is just a number.", "John Smith"); // Add a comment to A1 with text and author
oWorksheet.GetRange("A3").SetValue("Comment's author: "); // Set value "Comment's author:" in cell A3
oComment.SetAuthorName("Mark Potato"); // Change the author name of the comment
oWorksheet.GetRange("B3").SetValue(oComment.GetAuthorName()); // Set the value of B3 to the comment's author name
```