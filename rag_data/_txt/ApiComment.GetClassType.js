```vba
' This example gets a class type and inserts it into the table.
' Этот пример получает тип класса и вставляет его в таблицу.

Sub InsertClassType()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Get range A1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1")
    
    ' Add a comment to A1
    oRange.AddComment "This is just a number."
    
    ' Get the comment
    Dim oComment As Comment
    Set oComment = oRange.Comment
    
    ' Get the type of the comment
    Dim sType As String
    sType = TypeName(oComment)
    
    ' Insert the type into cell A3
    oWorksheet.Range("A3").Value = "Type: " & sType
End Sub
```

```javascript
// This example gets a class type and inserts it into the table.
// Этот пример получает тип класса и вставляет его в таблицу.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Get range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to A1
oRange.AddComment("This is just a number.");

// Get the comment
var oComment = oRange.GetComment();

// Get the class type of the comment
var sType = oComment.GetClassType();

// Insert the type into cell A3
oWorksheet.GetRange("A3").SetValue("Type: " + sType);
```