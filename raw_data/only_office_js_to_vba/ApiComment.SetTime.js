### Description / Описание

**English:**  
This code sets the value of cell A1 to "1", adds a comment "This is just a number." by "John Smith" to cell A1, and records the timestamp of the comment creation in cell B3.

**Русский:**  
Этот код устанавливает значение ячейки A1 на "1", добавляет комментарий "This is just a number." от "John Smith" в ячейку A1 и записывает временную метку создания комментария в ячейку B3.

---

#### Excel VBA Code

```vba
' VBA code to set cell values, add a comment, and record a timestamp

Sub AddCommentWithTimestamp()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    Dim currentTime As String
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the value of cell A1 to "1"
    ws.Range("A1").Value = "1"
    
    ' Set the range to cell A1
    Set rng = ws.Range("A1")
    
    ' Add a comment to cell A1
    Set cmt = rng.AddComment("This is just a number.")
    cmt.Author = "John Smith"
    
    ' Get the current timestamp
    currentTime = Format(Now, "mm/dd/yyyy hh:nn:ss")
    
    ' Set the value of cell A3 to "Timestamp: "
    ws.Range("A3").Value = "Timestamp: "
    
    ' Set the value of cell B3 to the current timestamp
    ws.Range("B3").Value = currentTime
End Sub
```

---

#### OnlyOffice JavaScript Code

```javascript
// JavaScript code to set cell values, add a comment, and record a timestamp

function addCommentWithTimestamp() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set the value of cell A1 to "1"
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get the range of cell A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to cell A1 with author "John Smith"
    var oComment = oRange.AddComment("This is just a number.", "John Smith");
    
    // Set the value of cell A3 to "Timestamp: "
    oWorksheet.GetRange("A3").SetValue("Timestamp: ");
    
    // Set the comment's timestamp to the current date and time
    oComment.SetTime(Date.now());
    
    // Set the value of cell B3 to the comment's timestamp
    oWorksheet.GetRange("B3").SetValue(oComment.GetTime());
}
```