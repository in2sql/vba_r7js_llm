**Description / Описание**

English: This script retrieves the active worksheet, sets the value of cell A1 to "1", adds a comment to cell A1, writes "Comment:" in cell A3, and places the comment ID in cell B3.

Russian: Этот скрипт получает активный лист, устанавливает значение ячейки A1 на "1", добавляет комментарий к ячейке A1, записывает "Comment:" в ячейку A3 и помещает ID комментария в ячейку B3.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Get the range for cell A1
var oRange = oWorksheet.GetRange("A1");

// Add a comment to cell A1
oRange.AddComment("This is just a number.");

// Set the value of cell A3 to "Comment: "
oWorksheet.GetRange("A3").SetValue("Comment: ");

// Set the value of cell B3 to the ID of the comment
oWorksheet.GetRange("B3").SetValue(oRange.GetComment().GetId());
```

```vba
' Excel VBA equivalent code

' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Set the value of cell A1 to "1"
oWorksheet.Range("A1").Value = "1"

' Add a comment to cell A1
With oWorksheet.Range("A1")
    .ClearComments ' Ensure no existing comment
    .AddComment "This is just a number."
End With

' Set the value of cell A3 to "Comment: "
oWorksheet.Range("A3").Value = "Comment: "

' Assuming there is a method to retrieve the comment ID
' Note: Excel VBA does not have a built-in Comment ID property
' This requires a custom implementation or additional API support
Dim commentID As String
commentID = GetCommentID(oWorksheet.Range("A1").Comment) ' Placeholder for comment ID retrieval
oWorksheet.Range("B3").Value = commentID

' Example function to get comment ID (requires custom implementation)
Function GetCommentID(cmt As Comment) As String
    ' Implementation to retrieve or generate a unique ID for the comment
    GetCommentID = "UniqueCommentID123" ' Replace with actual ID retrieval logic
End Function
```