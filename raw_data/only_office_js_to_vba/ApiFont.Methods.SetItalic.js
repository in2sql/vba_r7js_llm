## Description / Описание

**English:** This code sets the italic property to a specified portion of text in cell B1.

**Russian:** Этот код устанавливает свойство курсивного начертания для определенной части текста в ячейке B1.

```javascript
// This example sets the italic property to the specified font.
// Получение активного листа
var oWorksheet = Api.GetActiveSheet();
// Получение диапазона B1
var oRange = oWorksheet.GetRange("B1");
// Установка значения в ячейку B1
oRange.SetValue("This is just a sample text.");
// Получение символов с позиции 9 длиной 4
var oCharacters = oRange.GetCharacters(9, 4);
// Получение шрифта выбранных символов
var oFont = oCharacters.GetFont();
// Установка свойства курсивного начертания
oFont.SetItalic(true);
```

```vba
' This example sets the italic property to a specified portion of text in cell B1.
' Этот пример устанавливает свойство курсивного начертания для определенной части текста в ячейке B1.
Sub SetItalic()
    With Worksheets("Sheet1").Range("B1")
        ' Установка значения в ячейку B1
        .Value = "This is just a sample text."
        ' Установка свойства курсивного начертания для символов с позиции 9 длиной 4
        .Characters(Start:=9, Length:=4).Font.Italic = True
    End With
End Sub
```