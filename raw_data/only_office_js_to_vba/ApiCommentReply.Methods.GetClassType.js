# Code Description / Описание кода

**English:**  
This code retrieves the active worksheet, sets the value of cell A1 to "1", adds a comment to A1, replies to the comment, obtains the class type of the reply, and sets the value of cell A3 to display the reply's type.

**Русский:**  
Этот код получает активный лист, устанавливает значение ячейки A1 в "1", добавляет комментарий к A1, отвечает на комментарий, получает тип класса ответа и устанавливает значение ячейки A3 для отображения типа ответа.

```javascript
// JavaScript OnlyOffice API code
var oWorksheet = Api.GetActiveSheet(); // Retrieve the active worksheet
oWorksheet.GetRange("A1").SetValue("1"); // Set the value of cell A1 to "1"
var oRange = oWorksheet.GetRange("A1"); // Get the range for cell A1
var oComment = oRange.AddComment("This is just a number."); // Add a comment to cell A1
oComment.AddReply("Reply 1", "John Smith", "uid-1"); // Add a reply to the comment
var oReply = oComment.GetReply(); // Retrieve the reply
var sType = oReply.GetClassType(); // Get the class type of the reply
oWorksheet.GetRange("A3").SetValue("Type: " + sType); // Set the value of cell A3 to display the reply type
```

```vba
' Excel VBA code
Sub AddCommentAndReply()
    ' Retrieve the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the value of cell A1 to "1"
    oWorksheet.Range("A1").Value = "1"
    
    ' Add a comment to cell A1
    Dim oComment As Comment
    Set oComment = oWorksheet.Range("A1").AddComment("This is just a number.")
    
    ' Add a reply to the comment
    oComment.Replies.Add Text:="Reply 1", Author:="John Smith"
    
    ' Retrieve the reply
    Dim oReply As CommentThreadedReply
    Set oReply = oComment.Replies(1)
    
    ' Get the class type of the reply
    Dim sType As String
    sType = TypeName(oReply)
    
    ' Set the value of cell A3 to display the reply type
    oWorksheet.Range("A3").Value = "Type: " & sType
End Sub
```