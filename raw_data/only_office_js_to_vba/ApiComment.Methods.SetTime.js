```plaintext
// Description: This code sets the value of cell A1 to "1", adds a comment with the text "This is just a number." by "John Smith", sets the current timestamp in cell B3, and displays the timestamp.
// Описание: Этот код устанавливает значение ячейки A1 на "1", добавляет комментарий с текстом "Это просто число." от "John Smith", устанавливает текущую метку времени в ячейку B3 и отображает метку времени.
```

```vba
' VBA Code to set cell A1, add a comment, and set the timestamp in cell B3

Sub AddCommentWithTimestamp()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to cell A1
    Set oComment = oRange.AddComment("This is just a number.")
    oComment.Author = "John Smith"
    
    ' Set value of cell A3 to "Timestamp: "
    oWorksheet.Range("A3").Value = "Timestamp: "
    
    ' Set the timestamp of the comment
    oComment.Shape.TextFrame.Characters.Text = oComment.Text & vbLf & "Created on: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Set value of cell B3 to the timestamp
    oWorksheet.Range("B3").Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Sub
```

```javascript
// This example sets the timestamp of the comment creation in the current time zone format.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");

// Get range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1 with author "John Smith"
var oComment = oRange.AddComment("This is just a number.", "John Smith");

// Set value of cell A3 to "Timestamp: "
oWorksheet.GetRange("A3").SetValue("Timestamp: ");

// Set the timestamp of the comment to the current date and time
oComment.SetTime(Date.now());

// Set value of cell B3 to the comment's timestamp
oWorksheet.GetRange("B3").SetValue(oComment.GetTime());
```