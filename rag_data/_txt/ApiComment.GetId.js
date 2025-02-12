# Code Description / Описание кода

**English:** This code demonstrates how to set a value in cell A1, add a comment to it, retrieve the comment ID, and display the ID in cell B3.

**Russian:** Этот код демонстрирует, как установить значение в ячейку A1, добавить к ней комментарий, получить идентификатор комментария и отобразить идентификатор в ячейку B3.

```javascript
// This example shows how to get the comment ID.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("1");
var oRange = oWorksheet.GetRange("A1");
oRange.AddComment("This is just a number.");
oWorksheet.GetRange("A3").SetValue("Comment: ");
oWorksheet.GetRange("B3").SetValue(oRange.GetComment().GetId()); 
```

```vba
' This example shows how to get the comment ID.
Sub AddCommentAndGetID()
    Dim ws As Worksheet
    Dim rng As Range
    Dim commentID As String
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Set the value "1" in cell A1
    ws.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    ws.Range("A1").AddComment "This is just a number."
    
    ' Set the label in cell A3
    ws.Range("A3").Value = "Comment: "
    
    ' Retrieve the comment ID (Note: Excel VBA does not have a built-in Comment ID property)
    ' This is a placeholder as Excel VBA does not support retrieving a comment ID directly
    commentID = "N/A" ' Replace with appropriate logic if available
    
    ' Set the comment ID in cell B3
    ws.Range("B3").Value = commentID
End Sub
```