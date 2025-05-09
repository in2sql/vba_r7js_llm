# Code Description / Описание кода

**English:**  
This code demonstrates how to retrieve the parent `ApiCharacters` object of a specified font, set a value to a range, access specific characters, obtain the font properties, and modify the parent object's text.

**Russian:**  
Этот код демонстрирует, как получить объект `ApiCharacters` родительского шрифта, установить значение диапазона, получить определенные символы, получить свойства шрифта и изменить текст родительского объекта.

```javascript
// This code demonstrates how to retrieve the parent ApiCharacters object of a specified font

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range "B1"
var oRange = oWorksheet.GetRange("B1");

// Set the value of the range
oRange.SetValue("This is just a sample text.");

// Get characters starting at position 23 with length 4
var oCharacters = oRange.GetCharacters(23, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Get the parent object of the font
var oParent = oFont.GetParent();

// Set the text of the parent object
oParent.SetText("string");
```

```vba
' This code demonstrates how to retrieve the parent ApiCharacters object of a specified font

Sub ManipulateFont()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet
    
    ' Get the range "B1"
    Dim oRange As Object
    Set oRange = oWorksheet.GetRange("B1")
    
    ' Set the value of the range
    oRange.SetValue "This is just a sample text."
    
    ' Get characters starting at position 23 with length 4
    Dim oCharacters As Object
    Set oCharacters = oRange.GetCharacters(23, 4)
    
    ' Get the font of the selected characters
    Dim oFont As Object
    Set oFont = oCharacters.GetFont
    
    ' Get the parent object of the font
    Dim oParent As Object
    Set oParent = oFont.GetParent
    
    ' Set the text of the parent object
    oParent.SetText "string"
End Sub
```