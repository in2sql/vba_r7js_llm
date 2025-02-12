**Description / Описание**

This code sets the value of cell A1 to "1", adds a comment "This is just a number." to that cell, and then updates the comment text to "New comment text".

Этот код устанавливает значение ячейки A1 равным "1", добавляет комментарий "This is just a number." к этой ячейке, а затем обновляет текст комментария на "New comment text".

```vba
' VBA Code to set cell A1 value and update its comment

Sub SetCellValueAndComment()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the value of cell A1 to 1
    ws.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    Dim cmt As Comment
    On Error Resume Next
    Set cmt = ws.Range("A1").Comment
    On Error GoTo 0
    
    If cmt Is Nothing Then
        Set cmt = ws.Range("A1").AddComment("This is just a number.")
    End If
    
    ' Update the comment text
    cmt.Text Text:="New comment text"
End Sub
```

```javascript
// JavaScript Code to set cell A1 value and update its comment using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Update the comment text
oComment.SetText("New comment text");
```