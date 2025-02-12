**Description / Описание:**

This code creates a new history point, sets a value in cell A1, creates a history point, and then sets a value in cell A3.

Этот код создает новую точку истории, устанавливает значение в ячейку A1, создает точку истории и затем устанавливает значение в ячейку A3.

```javascript
// JavaScript code using OnlyOffice API
// This example creates a new history point.

var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue("This is just a sample text.");
Api.CreateNewHistoryPoint();
oWorksheet.GetRange("A3").SetValue("New history point was just created.");
```

```vba
' VBA code equivalent
' This example creates a new history point.

Sub CreateHistoryPoint()
    Dim oWorksheet As Worksheet
    ' Get the active sheet
    Set oWorksheet = ActiveSheet
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "This is just a sample text."
    ' Create new history point
    Call CreateNewHistoryPoint
    ' Set value in cell A3
    oWorksheet.Range("A3").Value = "New history point was just created."
End Sub

Sub CreateNewHistoryPoint()
    ' Placeholder for creating a history point in VBA
    ' Implement your own history tracking logic here
End Sub
```