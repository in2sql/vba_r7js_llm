# English Description
This code sets a value in cell B1, applies subscript formatting to a portion of the text, and displays whether subscript is applied in cell B3.

# Russian Description
Этот код устанавливает значение в ячейке B1, применяет форматирование подстрочником к части текста и отображает, применяется ли подстрочник, в ячейке B3.

```javascript
// This example shows how to get the subscript property of the specified font.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set value in B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters from position 9, length 4
var oFont = oCharacters.GetFont(); // Get the font of the characters
oFont.SetSubscript(true); // Set subscript to true
var bSubscript = oFont.GetSubscript(); // Get the subscript property
oWorksheet.GetRange("B3").SetValue("Subscript property: " + bSubscript); // Set value in B3
```

```vba
' This VBA code sets a value in cell B1, applies subscript formatting to a portion of the text,
' and displays whether subscript is applied in cell B3.

Sub SetSubscriptExample()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cellText As String
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value in B1
    Set rng = ws.Range("B1")
    rng.Value = "This is just a sample text."
    
    ' Apply subscript to characters 9 to 12 (4 characters starting at 9)
    With rng.Characters(Start:=9, Length:=4).Font
        .Subscript = True ' Set subscript to true
    End With
    
    ' Check if subscript is applied
    If rng.Characters(Start:=9, Length:=4).Font.Subscript = True Then
        cellText = "Subscript property: True"
    Else
        cellText = "Subscript property: False"
    End If
    
    ' Set value in B3
    ws.Range("B3").Value = cellText
End Sub
```