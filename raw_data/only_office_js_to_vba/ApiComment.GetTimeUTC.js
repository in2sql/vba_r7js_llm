```vba
' This example demonstrates how to retrieve the UTC timestamp of a comment's creation.
' Этот пример демонстрирует, как получить метку времени создания комментария в формате UTC.

Sub GetCommentTimestampUTC()
    Dim ws As Worksheet
    Dim rng As Range
    Dim commentText As String
    Dim commentTimestamp As Date
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value "1" to cell A1
    ws.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    ws.Range("A1").AddComment "This is just a number."
    
    ' Retrieve the comment from cell A1
    commentText = ws.Range("A1").Comment.Text
    
    ' Assuming the comment has a timestamp property (VBA does not support this directly)
    ' You might need to store the timestamp manually when adding the comment
    ' For demonstration, we'll use the current UTC time
    commentTimestamp = Now
    Debug.Print "Timestamp UTC: " & Format(commentTimestamp, "yyyy-mm-dd hh:nn:ss")
    
    ' Set the timestamp to cell B3
    ws.Range("A3").Value = "Timestamp UTC:"
    ws.Range("B3").Value = Format(commentTimestamp, "yyyy-mm-dd hh:nn:ss")
End Sub
```

```javascript
// This example shows how to get the timestamp of the comment creation in UTC format.
// Этот пример показывает, как получить метку времени создания комментария в UTC формате.

var oWorksheet = Api.GetActiveSheet();

// Set value "1" to cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get the range for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Set label for UTC timestamp in cell A3
oWorksheet.GetRange("A3").SetValue("Timestamp UTC: ");

// Set the UTC timestamp of the comment in cell B3
oWorksheet.GetRange("B3").SetValue(oComment.GetTimeUTC());
```