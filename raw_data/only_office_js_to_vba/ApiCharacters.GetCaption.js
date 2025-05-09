**Description / Описание**

English: This code sets the value of cell B1, extracts a substring from it starting at the 23rd character with a length of 4 characters, and writes a caption containing the extracted substring into cell B3.

Русский: Этот код устанавливает значение ячейки B1, извлекает подстроку, начиная с 23-го символа длиной 4 символа, и записывает подпись, содержащую извлеченную подстроку, в ячейку B3.

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1 and set its value
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");

// Get characters starting at position 23 with length 4
var oCharacters = oRange.GetCharacters(23, 4);

// Get the caption from the extracted characters
var sCaption = oCharacters.GetCaption();

// Set the caption value in cell B3
oWorksheet.GetRange("B3").SetValue("Caption: " + sCaption);
```

```vba
' VBA code equivalent to the OnlyOffice JavaScript example

Sub SetCaption()
    Dim ws As Worksheet
    Dim rngB1 As Range
    Dim rngB3 As Range
    Dim fullText As String
    Dim caption As String

    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set value in cell B1
    Set rngB1 = ws.Range("B1")
    rngB1.Value = "This is just a sample text."

    ' Get the full text from cell B1
    fullText = rngB1.Value

    ' Extract substring starting at 23rd character with length 4
    ' VBA uses 1-based indexing
    caption = Mid(fullText, 23, 4)

    ' Set the caption in cell B3
    Set rngB3 = ws.Range("B3")
    rngB3.Value = "Caption: " & caption
End Sub
```