# Description / Описание

This code checks if a comment is solved or not.
Этот код проверяет, решен ли комментарий.

```javascript
// This example checks if a comment is solved or not.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
var oComment = oRange.AddComment("This is just a number.");
oWorksheet.GetRange("A3").SetValue("Comment is solved: ");
oWorksheet.GetRange("B3").SetValue(oComment.IsSolved());
```

```vba
' This example checks if a comment is solved or not.
Sub CheckCommentSolved()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Set value in cell A3
    oWorksheet.Range("A3").Value = "Comment is solved: "
    
    ' Simulate IsSolved status
    ' VBA does not have an IsSolved method, so we use a custom approach
    ' For example, using a cell to store the solved status
    oWorksheet.Range("B3").Value = "False" ' Replace with actual logic as needed
End Sub
```