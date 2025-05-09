# Description
This code sets values in cells A1, B1, and C1, defines a named range "summa" referring to the sum of A1 and B1, and writes a description in cell A3.
Этот код устанавливает значения в ячейки A1, B1 и C1, определяет именованный диапазон "summa", ссылающийся на сумму A1 и B1, и устанавливает описательный текст в ячейку A3.

## Excel VBA Code
```vba
' This VBA code sets values in cells A1, B1, and C1, defines a named range "summa" referring to the sum of A1 and B1,
' and writes a description in cell A3.

Sub SetFormulaAndDefineName()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value in A1
    ws.Range("A1").Value = 1
    
    ' Set value in B1
    ws.Range("B1").Value = 2
    
    ' Set formula in C1
    ws.Range("C1").Formula = "=SUM(A1:B1)"
    
    ' Add defined name "summa" referring to A1:B1
    ThisWorkbook.Names.Add Name:="summa", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Set refers to formula for the defined name
    ThisWorkbook.Names("summa").RefersTo = "=SUM(A1:B1)"
    
    ' Set description in A3
    ws.Range("A3").Value = "The name 'summa' refers to the formula from the cell C1."
End Sub
```

## OnlyOffice JavaScript Code
```javascript
// This example sets a formula that the name is defined to refer to.
var oWorksheet = Api.GetActiveSheet();

// Set value in A1
oWorksheet.GetRange("A1").SetValue("1");

// Set value in B1
oWorksheet.GetRange("B1").SetValue("2");

// Set formula in C1
oWorksheet.GetRange("C1").SetValue("=SUM(A1:B1)");

// Add defined name "summa" referring to A1:B1
Api.AddDefName("summa", "Sheet1!$A$1:$B$1");

// Get the defined name "summa"
var oDefName = Api.GetDefName("summa");

// Set refers to formula for the defined name
oDefName.SetRefersTo("=SUM(A1:B1)");

// Set description in A3
oWorksheet.GetRange("A3").SetValue("The name 'summa' refers to the formula from the cell C1.");
```