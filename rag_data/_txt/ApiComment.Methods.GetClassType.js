**Description / Описание**

**English:**  
This code retrieves the active worksheet, sets the value of cell A1 to "1", adds a comment to A1, retrieves the comment's class type, and sets the value of cell A3 to display the class type.

**Русский:**  
Этот код получает активный лист, устанавливает значение ячейки A1 на "1", добавляет комментарий к A1, извлекает тип класса комментария и устанавливает значение ячейки A3 для отображения типа класса.

```javascript
// JavaScript OnlyOffice API code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range object for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
oRange.AddComment("This is just a number.");

// Retrieve the comment from cell A1
var oComment = oRange.GetComment();

// Get the class type of the comment
var sType = oComment.GetClassType();

// Set the value of cell A3 to display the class type
oWorksheet.GetRange("A3").SetValue("Type: " + sType);
```

```vba
' Excel VBA code equivalent

Sub InsertCommentAndType()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    oWorksheet.Range("A1").AddComment "This is just a number."
    
    ' Retrieve the comment from cell A1
    Dim oComment As Comment
    Set oComment = oWorksheet.Range("A1").Comment
    
    ' Get the class type of the comment
    Dim sType As String
    sType = TypeName(oComment)
    
    ' Set the value of cell A3 to display the class type
    oWorksheet.Range("A3").Value = "Type: " & sType
End Sub
```