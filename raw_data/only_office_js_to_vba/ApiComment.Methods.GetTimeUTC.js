---

**Description / Описание**

This code demonstrates how to add a value to cell A1, insert a comment with specific text, and retrieve the UTC timestamp of the comment creation, then display it in cell B3.

Этот код демонстрирует, как добавить значение в ячейку A1, вставить комментарий с определенным текстом и получить метку времени UTC создания комментария, а затем отобразить ее в ячейке B3.

---

**VBA Code / Код VBA**

```vba
Sub AddCommentWithTimestamp()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value "1" to cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to range A1
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Set label in cell A3
    oWorksheet.Range("A3").Value = "Timestamp UTC: "
    
    ' Set the UTC timestamp of the comment creation in cell B3
    oWorksheet.Range("B3").Value = Format(oComment.Date, "yyyy-mm-dd hh:nn:ss")
End Sub
```

---

**OnlyOffice JS Code / Код OnlyOffice JS**

```javascript
// This function adds a value to A1, inserts a comment, and retrieves the UTC timestamp of the comment creation.
function addCommentWithTimestamp() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value "1" to cell A1
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to range A1
    var oComment = oRange.AddComment("This is just a number.");
    
    // Set label in cell A3
    oWorksheet.GetRange("A3").SetValue("Timestamp UTC: ");
    
    // Set the UTC timestamp of the comment creation in cell B3
    oWorksheet.GetRange("B3").SetValue(oComment.GetTimeUTC());
}

// Execute the function
addCommentWithTimestamp();
```

---