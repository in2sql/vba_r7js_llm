```plaintext
// This script sets the value of cell A1 to "1", adds a comment to A1, deletes the comment, and writes a message to A3 indicating that the comment was deleted.
// Этот скрипт устанавливает значение ячейки A1 как "1", добавляет комментарий к A1, удаляет комментарий и записывает сообщение в A3, указывающее, что комментарий был удален.
```

```vba
' VBA Code to manipulate comments in Excel

Sub ManageComments()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the range A1 and assign value "1"
    Set rng = ws.Range("A1")
    rng.Value = "1"
    
    ' Add a comment to A1
    rng.AddComment "This is just a number."
    
    ' Get the comment from A1
    Set cmt = rng.Comment
    
    ' Delete the comment
    cmt.Delete
    
    ' Set value in A3 indicating the comment was deleted
    ws.Range("A3").Value = "The comment was just deleted from A1."
End Sub
```

```javascript
// This script sets the value of cell A1 to "1", adds a comment to A1, deletes the comment, and writes a message to A3 indicating that the comment was deleted.
// Этот скрипт устанавливает значение ячейки A1 как "1", добавляет комментарий к A1, удаляет комментарий и записывает сообщение в A3, указывающее, что комментарий был удален.

var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to A1
oRange.AddComment("This is just a number.");

// Get the comment from A1
var oComment = oRange.GetComment();

// Delete the comment
oComment.Delete();

// Set the value of cell A3 indicating the comment was deleted
oWorksheet.GetRange("A3").SetValue("The comment was just deleted from A1.");
```