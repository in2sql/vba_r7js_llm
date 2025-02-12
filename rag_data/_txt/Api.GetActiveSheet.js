// This script sets specific values and a formula in cells B1, B2, A3, and B3.
// Этот скрипт задает определенные значения и формулу в ячейках B1, B2, A3 и B3.

```vba
' Get the active sheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set value "2" to cell B1
oWorksheet.Range("B1").Value = "2"

' Set value "2" to cell B2
oWorksheet.Range("B2").Value = "2"

' Set value "2x2=" to cell A3
oWorksheet.Range("A3").Value = "2x2="

' Set formula "=B1*B2" to cell B3
oWorksheet.Range("B3").Formula = "=B1*B2"
```

```javascript
// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Set value "2" to cell B1
oWorksheet.GetRange("B1").SetValue("2");

// Set value "2" to cell B2
oWorksheet.GetRange("B2").SetValue("2");

// Set value "2x2=" to cell A3
oWorksheet.GetRange("A3").SetValue("2x2=");

// Set formula "=B1*B2" to cell B3
oWorksheet.GetRange("B3").SetValue("=B1*B2");
```