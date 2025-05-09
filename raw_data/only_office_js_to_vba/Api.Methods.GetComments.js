**Description / Описание**
This code adds comments to a worksheet, retrieves them, and displays the comment text and author in specific cells.
Этот код добавляет комментарии в электронную таблицу, извлекает их и отображает текст комментария и автора в определённых ячейках.

```vba
' VBA Code to add comments, retrieve them, and display in worksheet cells

Sub ManageComments()
    Dim Api As Object
    Dim oWorksheet As Worksheet
    Dim arrComments As Variant
    
    ' Initialize the OnlyOffice API object
    Set Api = CreateObject("OnlyOffice.Api")
    
    ' Add comments
    Api.AddComment "Comment 1", "Bob"
    Api.AddComment "Comment 2", "Bob"
    
    ' Retrieve comments
    Set arrComments = Api.GetComments()
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set values in cells A1 and B1 with comment text and author
    oWorksheet.Range("A1").Value = "Comment Text: " & arrComments(1).GetText()
    oWorksheet.Range("B1").Value = "Comment Author: " & arrComments(1).GetAuthorName()
End Sub
```

```javascript
// JavaScript Code to add comments, retrieve them, and display in worksheet cells

function manageComments() {
    // Initialize the OnlyOffice API
    var Api = new OnlyOffice.Api();
    
    // Add comments
    Api.AddComment("Comment 1", "Bob");
    Api.AddComment("Comment 2", "Bob");
    
    // Retrieve comments
    var arrComments = Api.GetComments();
    
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set values in cells A1 and B1 with comment text and author
    oWorksheet.GetRange("A1").SetValue("Comment Text: ", arrComments[0].GetText());
    oWorksheet.GetRange("B1").SetValue("Comment Author: ", arrComments[0].GetAuthorName());
}
```