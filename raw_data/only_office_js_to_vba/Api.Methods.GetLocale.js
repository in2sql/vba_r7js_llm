### Description / Описание

**English**: This code sets the locale to "en-CA", retrieves the current locale, and sets the value of cell A1 to display the locale ID.

**Russian**: Этот код устанавливает региональные настройки на "en-CA", получает текущий идентификатор локали и устанавливает значение ячейки A1 для отображения идентификатора локали.

#### VBA Code

```vba
' This code sets the locale to "en-CA", retrieves the current locale, and sets cell A1 with the locale ID.

Sub SetLocaleExample()
    Dim localeID As String
    ' Set the locale to Canadian English
    Application.LanguageSettings.LanguageID(msoLanguageIDUI) = msoLanguageIDEnglishCanada
    ' Get the current locale
    localeID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    ' Set the value of cell A1 to display the locale
    ThisWorkbook.ActiveSheet.Range("A1").Value = "Locale: " & localeID
End Sub
```

#### OnlyOffice JS Code

```javascript
// This code sets the locale to "en-CA", retrieves the current locale, and sets cell A1 with the locale ID.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
Api.SetLocale("en-CA"); // Set the locale to Canadian English
var nLocale = Api.GetLocale(); // Get the current locale
oWorksheet.GetRange("A1").SetValue("Locale: " + nLocale); // Set the value of cell A1
```