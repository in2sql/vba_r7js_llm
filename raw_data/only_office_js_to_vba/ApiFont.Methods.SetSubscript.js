## Description / Описание

**English:**  
This code sets the subscript property for a specific portion of text within cell B1 in the active worksheet.

**Russian:**  
Этот код устанавливает свойство подстрочного текста для определенной части текста в ячейке B1 на активном листе.

```vba
' VBA Code to set subscript for specific characters in cell B1

Sub SetSubscript()
    ' Get the active worksheet
    Dim oSheet As Worksheet
    Set oSheet = ActiveSheet
    
    ' Get the range B1
    Dim oRange As Range
    Set oRange = oSheet.Range("B1")
    
    ' Set the value of cell B1
    oRange.Value = "This is just a sample text."
    
    ' Set subscript for characters 9 to 12
    oRange.Characters(Start:=9, Length:=4).Font.Subscript = True
End Sub
```

```javascript
// JavaScript Code to set subscript for specific characters in cell B1

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of cell B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Set the subscript property to true
oFont.SetSubscript(true);
```