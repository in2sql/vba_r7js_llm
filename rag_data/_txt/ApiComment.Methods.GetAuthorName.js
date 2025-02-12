---
**Description:**
This example demonstrates how to add a comment to cell A1 in a worksheet and retrieve the comment author's name, displaying it in cell B3.

**Описание:**
Этот пример демонстрирует, как добавить комментарий к ячейке A1 на листе и получить имя автора комментария, отображая его в ячейке B3.
---

```javascript
// This example shows how to get the comment author's name.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set the value "1" in cell A1
var oRange = oWorksheet.GetRange("A1"); // Get the range for cell A1
var oComment = oRange.AddComment("This is just a number."); // Add a comment to cell A1
oWorksheet.GetRange("A3").SetValue("Comment's author: "); // Set label in cell A3
oWorksheet.GetRange("B3").SetValue(oComment.GetAuthorName()); // Set the comment author's name in cell B3
```

```vba
' This example demonstrates how to add a comment to cell A1 in a worksheet and retrieve the comment author's name, displaying it in cell B3.

Sub AddCommentAndGetAuthor()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ws.Range("A1").Value = "1" ' Set the value "1" in cell A1
    
    Dim comment As Comment
    Set comment = ws.Range("A1").AddComment("This is just a number.") ' Add a comment to cell A1
    
    ws.Range("A3").Value = "Comment's author: " ' Set label in cell A3
    ws.Range("B3").Value = comment.Author ' Set the comment author's name in cell B3
End Sub
```