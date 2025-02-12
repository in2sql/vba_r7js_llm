**Description / Описание**

This code demonstrates how to define a named formula in the active worksheet, set values in specific cells, and use the defined name in a formula.

Этот код демонстрирует, как определить именованную формулу в активном листе, установить значения в определенные ячейки и использовать определенное имя в формуле.

---

```javascript
// JavaScript code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set values in cells A1 and B1
oWorksheet.GetRange("A1").SetValue("1");
oWorksheet.GetRange("B1").SetValue("2");

// Set a formula in cell C1 that sums A1 and B1
oWorksheet.GetRange("C1").SetValue("=SUM(A1:B1)");

// Add a defined name 'summa' referring to the range A1:B1 on Sheet1
Api.AddDefName("summa", "Sheet1!$A$1:$B$1");

// Get the defined name 'summa' and set its reference to the sum formula
var oDefName = Api.GetDefName("summa");
oDefName.SetRefersTo("=SUM(A1:B1)");

// Set a descriptive text in cell A3
oWorksheet.GetRange("A3").SetValue("The name 'summa' refers to the formula from the cell C1.");

// Display the formula that 'summa' refers to in cell A4
oWorksheet.GetRange("A4").SetValue("Formula: " + oDefName.GetRefersTo());
```

```vba
' VBA equivalent code using OnlyOffice API

Sub DefineNamedFormula()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Set values in cells A1 and B1
    oWorksheet.GetRange("A1").SetValue "1"
    oWorksheet.GetRange("B1").SetValue "2"
    
    ' Set a formula in cell C1 that sums A1 and B1
    oWorksheet.GetRange("C1").SetValue "=SUM(A1:B1)"
    
    ' Add a defined name 'summa' referring to the range A1:B1 on Sheet1
    Api.AddDefName "summa", "Sheet1!$A$1:$B$1"
    
    ' Get the defined name 'summa' and set its reference to the sum formula
    Dim oDefName As Object
    Set oDefName = Api.GetDefName("summa")
    oDefName.SetRefersTo "=SUM(A1:B1)"
    
    ' Set a descriptive text in cell A3
    oWorksheet.GetRange("A3").SetValue "The name 'summa' refers to the formula from the cell C1."
    
    ' Display the formula that 'summa' refers to in cell A4
    oWorksheet.GetRange("A4").SetValue "Formula: " & oDefName.GetRefersTo()
End Sub
```