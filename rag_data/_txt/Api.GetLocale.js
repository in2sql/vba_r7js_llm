---

**Description / Описание:**

*This script demonstrates how to set and retrieve the locale in OnlyOffice using JavaScript and its equivalent in Excel VBA. The locale is updated to "en-CA" and the current locale ID is displayed in cell A1.*

*Этот скрипт демонстрирует, как установить и получить локаль в OnlyOffice с использованием JavaScript и его эквивалента в Excel VBA. Локаль обновляется до "en-CA", и текущий идентификатор локали отображается в ячейке A1.*

---

### OnlyOffice JavaScript

```javascript
// This example shows how to get the current locale ID.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
Api.SetLocale("en-CA"); // Set the locale to English (Canada)
var nLocale = Api.GetLocale(); // Retrieve the current locale ID
oWorksheet.GetRange("A1").SetValue("Locale: " + nLocale); // Set the value in cell A1
```

### Excel VBA

```vba
' This example shows how to set and get the current locale ID
Sub SetAndGetLocale()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet ' Get the active worksheet
    
    ' Set the locale to English (Canada) using the LocaleID 3084
    Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 3084
    
    Dim nLocale As Long
    nLocale = Application.LanguageSettings.LanguageID(msoLanguageIDUI) ' Retrieve the current locale ID
    
    oWorksheet.Range("A1").Value = "Locale: " & nLocale ' Set the value in cell A1
End Sub
```