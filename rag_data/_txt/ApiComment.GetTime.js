# Code Description / Описание кода

**English:** This code sets the value "1" in cell A1, adds a comment to it, and then writes the timestamp of the comment's creation in cell B3. It also sets the text "Timestamp:" in cell A3.

**Russian:** Этот код устанавливает значение "1" в ячейку A1, добавляет к ней комментарий, а затем записывает метку времени создания комментария в ячейку B3. Также устанавливает текст "Timestamp:" в ячейку A3.

```vba
' VBA Code
Sub AddCommentAndTimestamp()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    Dim commentTime As String
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value "1" in cell A1
    ws.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    Set rng = ws.Range("A1")
    Set cmt = rng.AddComment("This is just a number.")
    
    ' Store the current time as the comment's timestamp
    commentTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Set "Timestamp:" in cell A3
    ws.Range("A3").Value = "Timestamp:"
    
    ' Set the timestamp in cell B3
    ws.Range("B3").Value = commentTime
End Sub
```

```javascript
// JavaScript Code
// This example shows how to get the timestamp of the comment creation in the current time zone format.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
var oComment = oRange.AddComment("This is just a number.");
oWorksheet.GetRange("A3").SetValue("Timestamp: ");
oWorksheet.GetRange("B3").SetValue(oComment.GetTime()); 
```