### Description
This script sets a value in cell A1, adds a comment, deletes the comment, and sets a value in A3.
Этот скрипт устанавливает значение в ячейку A1, добавляет комментарий, удаляет комментарий и устанавливает значение в ячейку A3.

```vba
' Set A1 to "1"
Range("A1").Value = "1"

' Add comment to A1
Range("A1").AddComment "This is just a number."

' Delete the comment from A1
Range("A1").Comment.Delete

' Set A3 to indicate deletion
Range("A3").Value = "The comment was just deleted from A1."
```

```javascript
// This example deletes the ApiComment object.
var oWorksheet = Api.GetActiveSheet();

// Set A1 to "1"
oWorksheet.GetRange("A1").SetValue("1");

// Add comment to A1
var oRange = oWorksheet.GetRange("A1");
oRange.AddComment("This is just a number.");

// Delete the comment from A1
var oComment = oRange.GetComment();
oComment.Delete();

// Set A3 to indicate deletion
oWorksheet.GetRange("A3").SetValue("The comment was just deleted from A1.");
```