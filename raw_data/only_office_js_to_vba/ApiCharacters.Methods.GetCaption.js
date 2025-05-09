**Description / Описание**

*English:* This code sets the value of cell B1 to a sample text, extracts a substring starting from the 23rd character with a length of 4 characters, and sets cell B3 with the extracted caption.

*Russian:* Этот код устанавливает значение ячейки B1 на образец текста, извлекает подстроку начиная с 23-го символа длиной 4 символа и устанавливает ячейку B3 с извлеченной подписью.

```vba
' VBA Code equivalent to the OnlyOffice JS example

' Get the active sheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the range B1 and set its value
Dim oRange As Range
Set oRange = oWorksheet.Range("B1")
oRange.Value = "This is just a sample text."

' Extract characters starting at position 23 with length 4
Dim sText As String
sText = oRange.Value
Dim sCaption As String
sCaption = Mid(sText, 23, 4)

' Set the extracted caption to cell B3
oWorksheet.Range("B3").Value = "Caption: " & sCaption
```

```javascript
// OnlyOffice JS Code equivalent to the VBA example

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1 and set its value
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");

// Extract characters starting at position 23 with length 4
var oCharacters = oRange.GetCharacters(23, 4);
var sCaption = oCharacters.GetCaption();

// Set the extracted caption to cell B3
oWorksheet.GetRange("B3").SetValue("Caption: " + sCaption); 
```