**Description:**
This script sets the value "1" in cell A1, adds a comment to A1, writes "Timestamp:" in A3, and writes the timestamp of the comment's creation in B3.

**Описание:**
Этот скрипт устанавливает значение "1" в ячейке A1, добавляет комментарий к A1, записывает "Timestamp:" в A3 и записывает метку времени создания комментария в B3.

```javascript
// This script sets the value "1" in cell A1, adds a comment to A1, 
// writes "Timestamp:" in A3, and writes the timestamp of the comment's creation in B3.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1"); // Set value in A1
var oRange = oWorksheet.GetRange("A1"); 
var oComment = oRange.AddComment("This is just a number."); // Add comment to A1
oWorksheet.GetRange("A3").SetValue("Timestamp: "); // Set label in A3
oWorksheet.GetRange("B3").SetValue(oComment.GetTime()); // Set timestamp in B3
```

```vba
' This script sets the value "1" in cell A1, adds a comment to A1, 
' writes "Timestamp:" in A3, and writes the timestamp of the comment's creation in B3.

Sub AddCommentTimestamp()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    Set oWorksheet = ThisWorkbook.ActiveSheet ' Get active worksheet
    oWorksheet.Range("A1").Value = "1" ' Set value in A1
    Set oRange = oWorksheet.Range("A1")
    Set oComment = oRange.AddComment("This is just a number.") ' Add comment to A1
    oWorksheet.Range("A3").Value = "Timestamp: " ' Set label in A3
    oWorksheet.Range("B3").Value = oComment.Date ' Set timestamp in B3
End Sub
```