# This example shows how to get the quote text of the comment.
# Этот пример показывает, как получить текст цитаты комментария.

```vba
' This example shows how to get the quote text of the comment.
' Этот пример показывает, как получить текст цитаты комментария.

Sub GetCommentQuoteText()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oComment As Comment

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Set value in A1
    oWorksheet.Range("A1").Value = "1"

    ' Get range A1
    Set oRange = oWorksheet.Range("A1")

    ' Add comment to A1
    Set oComment = oRange.AddComment("This is just a number.")

    ' Set value in A3
    oWorksheet.Range("A3").Value = "Comment's quote text: "

    ' Set B3 to the quote text of the comment
    oWorksheet.Range("B3").Value = oComment.Text
End Sub
```

```javascript
// This example shows how to get the quote text of the comment.
// Этот пример показывает, как получить текст цитаты комментария.

function getCommentQuoteText() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();

    // Set value in A1
    oWorksheet.GetRange("A1").SetValue("1");

    // Get range A1
    var oRange = oWorksheet.GetRange("A1");

    // Add comment to A1
    var oComment = oRange.AddComment("This is just a number.");

    // Set value in A3
    oWorksheet.GetRange("A3").SetValue("Comment's quote text: ");

    // Set B3 to the quote text of the comment
    oWorksheet.GetRange("B3").SetValue(oComment.GetQuoteText());
}
```