# Get the User ID of the Comment Author / Получение ID пользователя автора комментария

This code demonstrates how to add a comment to a cell and retrieve the comment author's user ID.

Этот код демонстрирует, как добавить комментарий в ячейку и получить ID пользователя автора комментария.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range object for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Set the value of cell A3 to "Comment's user Id: "
oWorksheet.GetRange("A3").SetValue("Comment's user Id: ");

// Retrieve and set the comment author's user ID in cell B3
oWorksheet.GetRange("B3").SetValue(oComment.GetUserId());
```

```vba
' VBA code equivalent to the OnlyOffice API example

Sub GetCommentAuthorUserId()
    ' Declare worksheet variable
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    Dim oComment As Comment
    Set oComment = oWorksheet.Range("A1").AddComment("This is just a number.")
    
    ' Set the value of cell A3 to "Comment's user Id: "
    oWorksheet.Range("A3").Value = "Comment's user Id: "
    
    ' Retrieve and set the comment author's name in cell B3
    ' Note: VBA does not have a direct method to get user ID, so using Author as a placeholder
    oWorksheet.Range("B3").Value = oComment.Author
End Sub
```