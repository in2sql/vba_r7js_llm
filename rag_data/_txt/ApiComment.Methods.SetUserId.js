## Description / Описание

**English:**  
The code sets the value of cell A1 to "1", adds a comment by "John Smith" with the text "This is just a number.", assigns a user ID "uid-2" to the comment, and displays the user ID in cell B3.

**Русский:**  
Код устанавливает значение ячейки A1 равным "1", добавляет комментарий от "John Smith" с текстом "This is just a number.", присваивает идентификатор пользователя "uid-2" комментарий и отображает идентификатор пользователя в ячейке B3.

### JavaScript Code

```javascript
// This script sets the value of A1, adds a comment by "John Smith" with a user ID, and displays the user ID in B3.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
// Add a comment to A1 with author "John Smith"
var oComment = oRange.AddComment("This is just a number.", "John Smith");
oWorksheet.GetRange("A3").SetValue("Comment's user Id: ");
// Set the user ID for the comment
oComment.SetUserId("uid-2");
// Retrieve and display the user ID in B3
oWorksheet.GetRange("B3").SetValue(oComment.GetUserId());
```

### VBA Code

```vba
' This macro sets the value of A1, adds a comment by "John Smith" with a user ID, and displays the user ID in B3.
Sub AddCommentWithUserId()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the value of cell A1 to "1"
    ws.Range("A1").Value = "1"
    
    ' Add a comment to cell A1 with the text "This is just a number."
    ws.Range("A1").AddComment "This is just a number."
    
    ' Set the author of the comment to "John Smith"
    ws.Range("A1").Comment.Author = "John Smith"
    
    ' Assign a user ID to the comment using the Tag property
    ws.Range("A1").Comment.Tag = "uid-2"
    
    ' Set the value of cell A3 to label the user ID
    ws.Range("A3").Value = "Comment's user Id: "
    
    ' Retrieve the user ID from the comment and set it in cell B3
    ws.Range("B3").Value = ws.Range("A1").Comment.Tag
End Sub
```