# Get Comment Author's Name / Получение имени автора комментария

This code retrieves the active worksheet, sets the value of cell A1 to '1', adds a comment to A1, and then displays the comment author's name in cell B3.

Этот код получает активный лист, устанавливает значение ячейки A1 равным '1', добавляет комментарий к ячейке A1 и затем отображает имя автора комментария в ячейке B3.

```javascript
// This code retrieves the active worksheet, sets the value of cell A1 to '1',
// adds a comment to A1, and then displays the comment author's name in cell B3.

var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
var oComment = oRange.AddComment("This is just a number.");
oWorksheet.GetRange("A3").SetValue("Comment's author: ");
oWorksheet.GetRange("B3").SetValue(oComment.GetAuthorName());
```

```vba
' This code retrieves the active worksheet, sets the value of cell A1 to "1",
' adds a comment to A1, and then displays the comment author's name in cell B3.

Sub GetCommentAuthorName()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    Set oWorksheet = ActiveSheet
    oWorksheet.Range("A1").Value = "1"
    Set oRange = oWorksheet.Range("A1")
    oRange.AddComment "This is just a number."
    Set oComment = oRange.Comment
    oWorksheet.Range("A3").Value = "Comment's author: "
    oWorksheet.Range("B3").Value = oComment.Author
End Sub
```