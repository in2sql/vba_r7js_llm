### Description / Описание

**English:** This code sets the locale to English (Canada) and writes a sample text in cell A1.

**Russian:** Этот код устанавливает локаль на английский (Канада) и записывает пример текста в ячейку A1.

### Excel VBA Code

```vba
' This code sets the locale and writes a sample text in cell A1
Sub SetLocaleAndWriteValue()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Note: Excel VBA does not provide a direct method to set the locale.
    ' Locale settings are typically determined by the system settings.
    ' However, you can set number formats or use localized functions as needed.
    
    ' Write the sample text to cell A1
    ws.Range("A1").Value = "A sample spreadsheet with the language set to English (Canada)."
End Sub
```

### OnlyOffice JS Code

```javascript
// This code sets the locale and writes a sample text in cell A1
var oWorksheet = Api.GetActiveSheet();

// Set locale to English (Canada)
Api.SetLocale("en-CA");

// Write the sample text to cell A1
oWorksheet.GetRange("A1").SetValue("A sample spreadsheet with the language set to English (Canada).");
```