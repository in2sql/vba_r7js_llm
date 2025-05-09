# Description / Описание

This code formats a cell with a specific number format and sets its value accordingly.  
Этот код форматирует ячейку с определенным числовым форматом и устанавливает ее значение соответственно.

## VBA Code

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ActiveSheet

' Define the number format
Dim oFormat As String
oFormat = "$#,##0"

' Set the number format for range A1
oWorksheet.Range("A1").NumberFormat = oFormat

' Set the value for range A1
oWorksheet.Range("A1").Value = 123456
```

## OnlyOffice JS Code

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Define the number format and format the value
var oFormat = Api.Format("123456", "$#,##0");

// Set the formatted value to cell A1
oWorksheet.GetRange("A1").SetValue(oFormat);
```