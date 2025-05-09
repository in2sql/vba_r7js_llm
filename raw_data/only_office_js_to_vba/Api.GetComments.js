### Description / Описание

**English:**  
This code adds comments to the worksheet, retrieves them, and sets the values of cells A1 and B1 with the comment text and author, respectively.

**Русский:**  
Этот код добавляет комментарии на лист, извлекает их и устанавливает значения ячеек A1 и B1 с текстом комментария и автором соответственно.

---

#### VBA Code

```vba
' VBA code to add comments, retrieve them, and set cell values accordingly
Sub ManageComments()
    Dim ws As Worksheet
    Dim commentText As String
    Dim commentAuthor As String
    Dim commentCell As Range
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Add Comment 1 by Bob
    Set commentCell = ws.Range("A2")
    commentCell.AddComment "Comment 1"
    commentCell.Comment.Author = "Bob"
    
    ' Add Comment 2 by Bob
    Set commentCell = ws.Range("A3")
    commentCell.AddComment "Comment 2"
    commentCell.Comment.Author = "Bob"
    
    ' Retrieve the first comment
    If ws.Comments.Count > 0 Then
        commentText = ws.Comments(1).Text
        commentAuthor = ws.Comments(1).Author
        
        ' Set values in A1 and B1
        ws.Range("A1").Value = "Comment Text: " & commentText
        ws.Range("B1").Value = "Comment Author: " & commentAuthor
    End If
End Sub
```

---

#### JavaScript Code (OnlyOffice API)

```javascript
// JavaScript code to add comments, retrieve them, and set cell values accordingly

// Add Comment 1 by Bob
Api.AddComment("Comment 1", "Bob");

// Add Comment 2 by Bob
Api.AddComment("Comment 2", "Bob");

// Retrieve the array of comments
var arrComments = Api.GetComments();

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 with the first comment's text
oWorksheet.GetRange("A1").SetValue("Comment Text: ", arrComments[0].GetText());

// Set the value of cell B1 with the first comment's author
oWorksheet.GetRange("B1").SetValue("Comment Author: ", arrComments[0].GetAuthorName());
```