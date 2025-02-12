### Example: Accessing Characters and Font Objects / Пример: Доступ к объектам Characters и Font

This code accesses the active worksheet, selects cell B1, sets its value, retrieves a subset of characters, accesses their font, gets the parent object, and sets new text.

Этот код получает активный лист, выбирает ячейку B1, устанавливает её значение, извлекает подмножество символов, получает их шрифт, получает родительский объект и устанавливает новый текст.

```vba
' VBA code equivalent to the OnlyOffice JS example

Sub ModifyFontParent()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value of B1
    oRange.Value = "This is just a sample text."
    
    ' Get the characters from position 23 with length 4
    ' Note: VBA uses 1-based indexing
    Dim oCharacters As Characters
    Set oCharacters = oRange.Characters(Start:=23, Length:=4)
    
    ' Get the font of the specified characters
    Dim oFont As Font
    Set oFont = oCharacters.Font
    
    ' Get the parent of the font, which is the Characters object
    Dim oParent As Range
    Set oParent = oFont.Parent.Parent ' Parent of Font is Characters, parent of Characters is Range
    
    ' Set the text of the parent Range
    oParent.Value = "string"
End Sub
```

```javascript
// OnlyOffice JS code example

// This example shows how to get the parent ApiCharacters object of the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(23, 4);
var oFont = oCharacters.GetFont();
var oParent = oFont.GetParent();
oParent.SetText("string"); 
```