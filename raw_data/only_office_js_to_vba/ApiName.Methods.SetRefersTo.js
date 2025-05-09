# Code Description

**English:**  
This VBA and OnlyOffice JavaScript code sets values in cells A1 and B1, assigns a formula to C1, adds a defined name 'summa' referring first to the range A1:B1 and then updates it to refer to the sum of A1 and B1. It also sets a text in A3 indicating the reference of the defined name.

**Русский:**  
Этот код VBA и JavaScript для OnlyOffice устанавливает значения в ячейки A1 и B1, назначает формулу в C1, добавляет определенное имя 'summa', сначала ссылающееся на диапазон A1:B1, а затем обновляет его для ссылки на сумму A1 и B1. Также устанавливается текст в A3, указывающий на ссылку определенного имени.

## VBA Code

```vba
' Set values to cells A1 and B1, assign formulas, and define a named range
Sub SetValuesAndName()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ws.Range("A1").Value = 1 ' Set A1 to 1
    ws.Range("B1").Value = 2 ' Set B1 to 2
    
    ws.Range("C1").Formula = "=SUM(A1:B1)" ' Set formula in C1
    
    ' Add defined name 'summa' referring to range A1:B1
    ThisWorkbook.Names.Add Name:="summa", RefersTo:="=Sheet1!$A$1:$B$1"
    
    ' Update defined name 'summa' to refer to the sum formula
    ThisWorkbook.Names("summa").RefersTo = "=SUM(A1:B1)"
    
    ws.Range("A3").Value = "The name 'summa' refers to the formula from the cell C1." ' Set text in A3
End Sub
```

## OnlyOffice JavaScript API Code

```javascript
// Set values to cells A1 and B1, assign formulas, and define a named range
function setValuesAndName() {
    var oWorksheet = Api.GetActiveSheet(); // Get active sheet
    
    oWorksheet.GetRange("A1").SetValue("1"); // Set A1 to 1
    oWorksheet.GetRange("B1").SetValue("2"); // Set B1 to 2
    
    oWorksheet.GetRange("C1").SetValue("=SUM(A1:B1)"); // Set formula in C1
    
    Api.AddDefName("summa", "Sheet1!$A$1:$B$1"); // Add defined name 'summa' referring to A1:B1
    var oDefName = Api.GetDefName("summa"); // Get defined name 'summa'
    oDefName.SetRefersTo("=SUM(A1:B1)"); // Update 'summa' to refer to the sum formula
    
    oWorksheet.GetRange("A3").SetValue("The name 'summa' refers to the formula from the cell C1."); // Set text in A3
}
```