**Description / Описание**

English: This code gets the active worksheet, sets a value in cell A1, adds a comment to that cell with a reply, retrieves the class type of the reply, and inserts the type into cell A3.

Russian: Этот код получает активный лист, устанавливает значение в ячейку A1, добавляет к этой ячейке комментарий с ответом, получает тип класса ответа и вставляет тип в ячейку A3.

**VBA Code:**
```vba
' This example gets a class type and inserts it into the table.
Sub InsertClassType()
    Dim oWorksheet As Object
    Dim oRange As Object
    Dim oComment As Object
    Dim oReply As Object
    Dim sType As String
    
    ' Get the active worksheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Set value "1" in cell A1
    oWorksheet.GetRange("A1").SetValue "1"
    
    ' Get range A1
    Set oRange = oWorksheet.GetRange("A1")
    
    ' Add a comment to cell A1
    Set oComment = oRange.AddComment("This is just a number.")
    
    ' Add a reply to the comment
    oComment.AddReply "Reply 1", "John Smith", "uid-1"
    
    ' Get the reply from the comment
    Set oReply = oComment.GetReply()
    
    ' Get the class type of the reply
    sType = oReply.GetClassType()
    
    ' Insert the type into cell A3
    oWorksheet.GetRange("A3").SetValue "Type: " & sType
End Sub
```

**JavaScript Code:**
```javascript
// This example gets a class type and inserts it into the table.
function insertClassType() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value "1" in cell A1
    oWorksheet.GetRange("A1").SetValue("1");
    
    // Get range A1
    var oRange = oWorksheet.GetRange("A1");
    
    // Add a comment to cell A1
    var oComment = oRange.AddComment("This is just a number.");
    
    // Add a reply to the comment
    oComment.AddReply("Reply 1", "John Smith", "uid-1");
    
    // Get the reply from the comment
    var oReply = oComment.GetReply();
    
    // Get the class type of the reply
    var sType = oReply.GetClassType();
    
    // Insert the type into cell A3
    oWorksheet.GetRange("A3").SetValue("Type: " + sType);
}
```