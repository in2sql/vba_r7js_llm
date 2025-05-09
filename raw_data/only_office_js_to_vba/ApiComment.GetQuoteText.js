# Code Description / Описание кода

**English:**  
This code sets the value "1" in cell A1, adds a comment "This is just a number." to A1, sets the value "Comment's quote text:" in cell A3, and then retrieves and sets the comment's quote text in cell B3.

**Russian:**  
Этот код устанавливает значение "1" в ячейку A1, добавляет комментарий "This is just a number." в ячейку A1, устанавливает значение "Comment's quote text:" в ячейку A3, а затем извлекает и устанавливает текст цитаты комментария в ячейку B3.

```vba
' VBA code to add a comment and retrieve its text

Sub AddCommentAndQuoteText()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set value "1" in cell A1
    ws.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    ws.Range("A1").AddComment "This is just a number."
    
    ' Set value "Comment's quote text:" in cell A3
    ws.Range("A3").Value = "Comment's quote text: "
    
    ' Retrieve the comment's text
    Dim commentText As String
    commentText = ws.Range("A1").Comment.Text
    
    ' Set the comment's text in cell B3
    ws.Range("B3").Value = commentText
End Sub
```

```javascript
// JavaScript code using OnlyOffice API to add a comment and retrieve its quote text

var oWorksheet = Api.GetActiveSheet();

// Set the value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Set the value "Comment's quote text:" in cell A3
oWorksheet.GetRange("A3").SetValue("Comment's quote text: ");

// Retrieve the comment's quote text and set it in cell B3
oWorksheet.GetRange("B3").SetValue(oComment.GetQuoteText());
```