**Description / Описание**

English: This code demonstrates how to add a comment to a cell and retrieve the comment text in a spreadsheet using OnlyOffice API.

Russian: Этот код демонстрирует, как добавить комментарий к ячейке и получить текст комментария в таблице, используя OnlyOffice API.

```vba
' VBA Code to add a comment to a cell and retrieve the comment text

Sub AddAndRetrieveComment()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    oWorksheet.Range("A1").AddComment "This is just a number."
    
    ' Set value in cell A3
    oWorksheet.Range("A3").Value = "Comment: "
    
    ' Retrieve the comment text from cell A1 and set it in cell B3
    oWorksheet.Range("B3").Value = oWorksheet.Range("A1").Comment.Text
End Sub
```

```javascript
// JavaScript Code to add a comment to a cell and retrieve the comment text

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get the range for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
oRange.AddComment("This is just a number.");

// Set value in cell A3
oWorksheet.GetRange("A3").SetValue("Comment: ");

// Retrieve the comment text from cell A1 and set it in cell B3
oWorksheet.GetRange("B3").SetValue(oRange.GetComment().GetText());
```