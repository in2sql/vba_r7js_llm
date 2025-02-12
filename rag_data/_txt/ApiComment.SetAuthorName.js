**Description / Описание**

This code sets the value of cell A1 to "1", adds a comment to it authored by "John Smith", changes the comment's author to "Mark Potato", and then displays the author's name in cell B3.

Этот код устанавливает значение ячейки A1 на "1", добавляет к ней комментарий, автором которого является "John Smith", изменяет автора комментария на "Mark Potato" и затем отображает имя автора в ячейке B3.

```vba
' VBA Code to manipulate comments in Excel

Sub SetCommentAuthor()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value of A1 to "1"
    ws.Range("A1").Value = "1"
    
    ' Get the range A1
    Set rng = ws.Range("A1")
    
    ' Add a comment to A1 with initial author "John Smith"
    Set cmt = rng.AddComment("This is just a number.")
    cmt.Author = "John Smith"
    
    ' Set value of A3
    ws.Range("A3").Value = "Comment's author: "
    
    ' Change the comment's author to "Mark Potato"
    cmt.Author = "Mark Potato"
    
    ' Set value of B3 to the comment's author name
    ws.Range("B3").Value = cmt.Author
End Sub
```

```javascript
// JavaScript Code to manipulate comments in OnlyOffice

// This function sets the comment author's name and updates cell values
function setCommentAuthor() {
    // Get the active sheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value of A1 to "1"
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get the range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to A1 with initial author "John Smith"
    var oComment = oRange.AddComment("This is just a number.", "John Smith");
    
    // Set value of A3
    oWorksheet.GetRange("A3").SetValue("Comment's author: ");
    
    // Change the comment's author to "Mark Potato"
    oComment.SetAuthorName("Mark Potato");
    
    // Set value of B3 to the comment's author name
    oWorksheet.GetRange("B3").SetValue(oComment.GetAuthorName());
}
```