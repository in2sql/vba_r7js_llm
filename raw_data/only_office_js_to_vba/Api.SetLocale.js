# Set Locale and Assign Value to Cell A1
# Установить локаль и присвоить значение ячейке A1

```vba
' This macro sets the locale to English (Canada) and assigns a value to cell A1
Sub SetLocaleAndAssignValue()
    ' Set the locale to English (Canada)
    ' Note: VBA does not provide a direct method to set locale for the workbook.
    ' Locale settings are typically managed at the application or system level.
    
    ' Assign value to cell A1
    With ThisWorkbook.ActiveSheet.Range("A1")
        .Value = "A sample spreadsheet with the language set to English (Canada)."
    End With
End Sub
```

```javascript
// This example sets a locale to the document and assigns a value to cell A1.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
Api.SetLocale("en-CA"); // Set the locale to English (Canada)
var oRange = oWorksheet.GetRange("A1"); // Get the range A1
oRange.SetValue("A sample spreadsheet with the language set to English (Canada)."); // Set value in A1
```