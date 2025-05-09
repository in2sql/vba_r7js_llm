### Description / Описание

**English:** This VBA script adds comments to the active Excel worksheet, retrieves them, and displays the first comment's text and author in specified cells.

**Русский:** Этот VBA-скрипт добавляет комментарии в активный лист Excel, извлекает их и отображает текст и автора первого комментария в указанных ячейках.

```vba
' Adds comments to the active worksheet and displays the first comment's text and author
Sub AddAndDisplayComments()
    ' Add the first comment with author "Bob" to cell A1
    With ActiveSheet.Range("A1")
        .ClearComments ' Remove existing comments
        .AddComment "Comment 1"
        .Comment.Author = "Bob"
    End With
    
    ' Add the second comment without specifying an author to cell A2
    With ActiveSheet.Range("A2")
        .ClearComments ' Remove existing comments
        .AddComment "Comment 2"
        ' Author will default to the application's author
    End With
    
    ' Retrieve comments from the active worksheet
    Dim cmt As Comment
    If ActiveSheet.Comments.Count > 0 Then
        Set cmt = ActiveSheet.Comments(1) ' Get the first comment
        
        ' Display comment text in cell A1
        ActiveSheet.Range("C1").Value = "Comment Text: " & cmt.Text
        
        ' Display comment author in cell B1
        ActiveSheet.Range("D1").Value = "Comment Author: " & cmt.Author
    End If
End Sub
```

---

### Description / Описание

**English:** This JavaScript code for OnlyOffice adds comments to the document, retrieves them, and displays the first comment's text and author in specified cells of the active worksheet.

**Русский:** Этот JavaScript-код для OnlyOffice добавляет комментарии в документ, извлекает их и отображает текст и автора первого комментария в указанных ячейках активного листа.

```javascript
// Adds comments to the OnlyOffice document and displays the first comment's text and author in the active sheet
Api.AddComment("Comment 1", "Bob"); // Add first comment with author "Bob"
Api.AddComment("Comment 2"); // Add second comment without specifying author
var arrComments = Api.GetComments(); // Retrieve all comments
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet

// Check if there are any comments
if (arrComments.length > 0) {
    // Set value in cell C1 with comment text
    oWorksheet.GetRange("C1").SetValue("Comment Text: " + arrComments[0].GetText());
    
    // Set value in cell D1 with comment author
    oWorksheet.GetRange("D1").SetValue("Comment Author: " + arrComments[0].GetAuthorName());
}
```