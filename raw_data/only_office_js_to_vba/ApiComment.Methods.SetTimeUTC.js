### Description / Описание

**English:** This code sets a value in cell A1, adds a comment to it with the author's name, records the current UTC timestamp, and displays the timestamp in cell B3.

**Русский:** Этот код устанавливает значение в ячейке A1, добавляет к ней комментарий с именем автора, записывает текущую временную метку UTC и отображает временную метку в ячейке B3.

```javascript
// This example sets the timestamp of the comment creation in UTC format.
var oWorksheet = Api.GetActiveSheet();

// Set the value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get the range object for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1 with the text and author name
var oComment = oRange.AddComment("This is just a number.", "John Smith");

// Set the label in cell A3
oWorksheet.GetRange("A3").SetValue("Timestamp UTC: ");

// Set the current UTC timestamp in the comment
oComment.SetTimeUTC(Date.now());

// Display the UTC timestamp in cell B3
oWorksheet.GetRange("B3").SetValue(oComment.GetTimeUTC());
```

```vba
' This VBA macro sets a value in cell A1, adds a comment with the author's name,
' records the current UTC timestamp, and displays the timestamp in cell B3.

Sub AddCommentWithUTCTimestamp()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cmt As Comment
    Dim currentTimeUTC As String
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the value "1" in cell A1
    ws.Range("A1").Value = "1"
    
    ' Set the range to cell A1
    Set rng = ws.Range("A1")
    
    ' Add a comment to cell A1 with the text and author name
    ' Remove existing comment if any
    On Error Resume Next
    rng.ClearComments
    On Error GoTo 0
    Set cmt = rng.AddComment("This is just a number.")
    cmt.Author = "John Smith"
    
    ' Set the label in cell A3
    ws.Range("A3").Value = "Timestamp UTC: "
    
    ' Get the current UTC time
    currentTimeUTC = Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")
    
    ' Add the UTC timestamp to the comment
    cmt.Text Text:=cmt.Text & vbCrLf & "Timestamp UTC: " & currentTimeUTC
    
    ' Display the UTC timestamp in cell B3
    ws.Range("B3").Value = currentTimeUTC
End Sub
```