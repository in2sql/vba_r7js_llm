**Description / Описание**

This code retrieves the active worksheet, formats the number "123456" according to the specified format, and sets the formatted value to cell A1.

Этот код получает активный лист, форматирует число "123456" согласно указанному формату и устанавливает отформатированное значение в ячейку A1.

---

```vba
' VBA Code to format a number and set it in cell A1

Sub FormatAndSetValue()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oFormat As String
    Dim value As Double
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the format
    oFormat = "$#,##0"
    
    ' Define the value
    value = 123456
    
    ' Get the range A1
    Set oRange = oWorksheet.Range("A1")
    
    ' Apply the format to the range
    oRange.NumberFormat = oFormat
    
    ' Set the value to the range
    oRange.Value = value
End Sub
```

```javascript
// JavaScript Code to format a number and set it in cell A1 using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Define the format expression
var oFormat = Api.Format("123456", "$#,##0");

// Get the range A1 and set the formatted value
oWorksheet.GetRange("A1").SetValue(oFormat);
```