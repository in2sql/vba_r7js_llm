# Code Description / Описание кода

**English:**  
This code sets the value of cell A1 to "1", adds a comment to it with a reply from "John Smith", and displays the reply text in cell B3.

**Русский:**  
Этот код устанавливает значение в ячейке A1 равным "1", добавляет к ней комментарий с ответом от "John Smith" и отображает текст ответа в ячейке B3.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
var oComment = oRange.AddComment("This is just a number.");

// Add a reply to the comment
oComment.AddReply("Reply 1", "John Smith", "uid-1");

// Retrieve the reply
var oReply = oComment.GetReply();

// Set the value of cell A3 to label the reply text
oWorksheet.GetRange("A3").SetValue("Comment's reply text: ");

// Set the value of cell B3 to the reply text
oWorksheet.GetRange("B3").SetValue(oReply.GetText());
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set the value of cell A1 to "1"
oWorksheet.Range("A1").Value = "1"

' Add a comment to cell A1
Dim oComment As Comment
Set oComment = oWorksheet.Range("A1").AddComment("This is just a number.")

' Add a reply to the comment
oComment.Replies.Add "Reply 1", "John Smith"

' Retrieve the reply text
Dim oReplyText As String
oReplyText = oComment.Replies(1).Text

' Set the value of cell A3 to label the reply text
oWorksheet.Range("A3").Value = "Comment's reply text: "

' Set the value of cell B3 to the reply text
oWorksheet.Range("B3").Value = oReplyText
```