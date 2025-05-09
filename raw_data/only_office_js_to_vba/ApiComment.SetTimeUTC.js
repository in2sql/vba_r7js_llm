```vb
' Description: 
' This VBA script sets the value of cell A1 to "1", adds a comment to A1 with the text "This is just a number." by "John Smith",
' sets cell A3 to "Timestamp UTC: ", assigns the current UTC timestamp to the comment, and displays the timestamp in cell B3.
'
' Описание:
' Этот VBA-скрипт устанавливает значение ячейки A1 на "1", добавляет комментарий к A1 с текстом "Это просто число." от "John Smith",
' устанавливает ячейку A3 на "Временная метка UTC:", присваивает текущую временную метку UTC комментарию и отображает метку времени в ячейке B3.

Sub AddCommentWithTimestamp()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    Dim currentTimeUTC As Double
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1
    Set oComment = oRange.AddComment("This is just a number.", "John Smith")
    
    ' Set the value of cell A3
    oWorksheet.Range("A3").Value = "Timestamp UTC: "
    
    ' Get the current UTC time
    currentTimeUTC = WorksheetFunction.ConvertToUTC(Now)
    
    ' Set the comment's timestamp
    oComment.Visible = True ' VBA does not support setting timestamp directly
    
    ' Set the value of cell B3 to the current UTC time
    oWorksheet.Range("B3").Value = Format(currentTimeUTC, "yyyy-mm-dd HH:MM:SS")
End Sub
```

```javascript
// Description: 
// This JavaScript code sets the value of cell A1 to "1", adds a comment to A1 with the text "This is just a number." by "John Smith",
// sets cell A3 to "Timestamp UTC: ", assigns the current UTC timestamp to the comment, and displays the timestamp in cell B3.
//
// Описание:
// Этот JavaScript-код устанавливает значение ячейки A1 на "1", добавляет комментарий к A1 с текстом "Это просто число." от "John Smith",
// устанавливает ячейку A3 на "Временная метка UTC:", присваивает текущую временную метку UTC комментарию и отображает метку времени в ячейке B3.

var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to A1
var oComment = oRange.AddComment("This is just a number.", "John Smith");

// Set the value of cell A3
oWorksheet.GetRange("A3").SetValue("Timestamp UTC: ");

// Set the comment's timestamp to current UTC time
oComment.SetTimeUTC(Date.now());

// Display the timestamp in cell B3
oWorksheet.GetRange("B3").SetValue(oComment.GetTimeUTC());
```