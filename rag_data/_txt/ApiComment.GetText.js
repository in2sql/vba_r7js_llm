**Description / Описание:**
This code demonstrates how to add a comment to a cell and retrieve its text in both OnlyOffice JavaScript API and Excel VBA.
Этот код демонстрирует, как добавить комментарий к ячейке и получить его текст как в OnlyOffice JavaScript API, так и в Excel VBA.

---

```javascript
// OnlyOffice JavaScript API Example

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get the range for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
oRange.AddComment("This is just a number.");

// Set the label "Comment:" in cell A3
oWorksheet.GetRange("A3").SetValue("Comment: ");

// Retrieve the comment text from cell A1 and set it in cell B3
oWorksheet.GetRange("B3").SetValue(oRange.GetComment().GetText());
```

```vba
' Excel VBA Equivalent Example

Sub AddAndRetrieveComment()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set the value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    oWorksheet.Range("A1").AddComment "This is just a number."
    
    ' Set the label "Comment:" in cell A3
    oWorksheet.Range("A3").Value = "Comment: "
    
    ' Retrieve the comment text from cell A1 and set it in cell B3
    On Error Resume Next ' In case there is no comment
    oWorksheet.Range("B3").Value = oWorksheet.Range("A1").Comment.Text
    On Error GoTo 0
End Sub
```