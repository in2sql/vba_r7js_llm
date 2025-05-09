**English:** This code marks a comment as solved in the active worksheet.

**Russian:** Этот код помечает комментарий как решённый в активном листе.

```javascript
// This example marks a comment as solved.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set the value of cell A1 to "1"
var oRange = oWorksheet.GetRange("A1"); // Get the range A1
var oComment = oRange.AddComment("This is just a number.", "John Smith"); // Add a comment to A1
oWorksheet.GetRange("A3").SetValue("Comment is solved: "); // Set the value of cell A3
oComment.SetSolved(true); // Mark the comment as solved
oWorksheet.GetRange("B3").SetValue(oComment.IsSolved()); // Set the value of cell B3 to the solved status
```

```vba
' This example marks a comment as solved.
Sub MarkCommentAsSolved()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    Set oWorksheet = ActiveSheet ' Get the active worksheet
    oWorksheet.Range("A1").Value = "1" ' Set the value of cell A1 to "1"
    
    Set oRange = oWorksheet.Range("A1") ' Get the range A1
    Set oComment = oRange.AddComment("This is just a number.") ' Add a comment to A1
    oComment.Author = "John Smith" ' Set the author of the comment
    
    oWorksheet.Range("A3").Value = "Comment is solved: " ' Set the value of cell A3
    oComment.Visible = False ' Typically, VBA does not have a direct 'SetSolved' method
    ' To simulate 'solved', you might change the comment's text or formatting
    oComment.Text oComment.Text & " [Solved]" ' Append [Solved] to the comment text
    
    oWorksheet.Range("B3").Value = "True" ' Set the value of cell B3 to indicate the comment is solved
End Sub
```