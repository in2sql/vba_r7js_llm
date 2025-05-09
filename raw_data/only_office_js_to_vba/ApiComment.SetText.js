# Description / Описание

**English:**  
This code sets the value of cell A1 to "1", adds a comment "This is just a number." to cell A1, and then updates the comment text to "New comment text".

**Russian:**  
Этот код устанавливает значение ячейки A1 на "1", добавляет комментарий "This is just a number." к ячейке A1, а затем обновляет текст комментария на "New comment text".

```vba
' VBA code
' This code sets the value of cell A1 to "1", adds a comment, and updates the comment text

Sub SetComment()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    
    Set ws = ActiveSheet
    Set rng = ws.Range("A1")
    
    rng.Value = "1" ' Set cell A1 value to "1"
    Set cmt = rng.AddComment("This is just a number.") ' Add initial comment
    cmt.Text Text:="New comment text" ' Update comment text
End Sub
```

```javascript
// JavaScript code
// This code sets the value of cell A1 to "1", adds a comment, and updates the comment text

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set cell A1 value to "1"
var oRange = oWorksheet.GetRange("A1"); // Get range A1
var oComment = oRange.AddComment("This is just a number."); // Add initial comment
oComment.SetText("New comment text"); // Update comment text
```