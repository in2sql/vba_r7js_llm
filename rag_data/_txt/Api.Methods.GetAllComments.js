### This code adds comments to the active worksheet and retrieves comment information.
### Этот код добавляет комментарии на активный лист и извлекает информацию о комментариях.

```vba
' VBA code to add comments and retrieve comment information

Sub ManageComments()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = Api.GetActiveSheet
    
    ' Add a comment to the worksheet
    Api.AddComment "Comment 1", "John Smith"
    
    ' Add a comment to cell A4
    oWorksheet.Range("A4").AddComment "Comment 2", "Mark Potato"
    
    ' Retrieve all comments
    Dim arrComments As Collection
    Set arrComments = Api.GetAllComments
    
    ' Set value in cell A1 with the text of the second comment
    oWorksheet.Range("A1").Value = "Comment text: " & arrComments(1).GetText()
    
    ' Set value in cell A2 with the author of the second comment
    oWorksheet.Range("A2").Value = "Comment author: " & arrComments(1).GetAuthorName()
End Sub
```

```javascript
// JavaScript code to add comments and retrieve comment information

var oWorksheet = Api.GetActiveSheet();
Api.AddComment("Comment 1", "John Smith");
oWorksheet.GetRange("A4").AddComment("Comment 2", "Mark Potato");
var arrComments = Api.GetAllComments();
oWorksheet.GetRange("A1").SetValue("Comment text: " + arrComments[1].GetText());
oWorksheet.GetRange("A2").SetValue("Comment author: " + arrComments[1].GetAuthorName());
```