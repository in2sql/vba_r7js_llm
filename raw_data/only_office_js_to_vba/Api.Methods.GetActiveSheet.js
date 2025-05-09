# Description / Описание

**English:**  
This code retrieves the active worksheet, sets the values "2" in cells B1 and B2, sets "2x2=" in cell A3, and assigns the formula `=B1*B2` to cell B3.

**Russian:**  
Этот код получает активный лист, устанавливает значения "2" в ячейках B1 и B2, устанавливает "2x2=" в ячейке A3 и присваивает формулу `=B1*B2` ячейке B3.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value "2" in cell B1
oWorksheet.GetRange("B1").SetValue("2");

// Set value "2" in cell B2
oWorksheet.GetRange("B2").SetValue("2");

// Set value "2x2=" in cell A3
oWorksheet.GetRange("A3").SetValue("2x2=");

// Assign formula "=B1*B2" to cell B3
oWorksheet.GetRange("B3").SetValue("=B1*B2");
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Set value "2" in cell B1
oWorksheet.Range("B1").Value = "2"

' Set value "2" in cell B2
oWorksheet.Range("B2").Value = "2"

' Set value "2x2=" in cell A3
oWorksheet.Range("A3").Value = "2x2="

' Assign formula "=B1*B2" to cell B3
oWorksheet.Range("B3").Formula = "=B1*B2"
```