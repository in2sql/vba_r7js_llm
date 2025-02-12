# Description / Описание

This code demonstrates how to define and manipulate named ranges and formulas in Excel using both VBA and OnlyOffice JavaScript APIs.

Этот код демонстрирует, как определить и управлять именованными диапазонами и формулами в Excel с использованием как VBA, так и API OnlyOffice на JavaScript.

```vba
' VBA Code
Sub DefineAndManipulateNamedRange()
    Dim oWorksheet As Worksheet
    Dim oDefName As Name
    
    ' Set the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set values in cells A1 and B1
    oWorksheet.Range("A1").Value = 1
    oWorksheet.Range("B1").Value = 2
    
    ' Set formula in cell C1
    oWorksheet.Range("C1").Formula = "=SUM(A1:B1)"
    
    ' Add a defined name "summa" referring to the range A1:B1
    Set oDefName = ThisWorkbook.Names.Add(Name:="summa", RefersTo:="=SUM(A1:B1)")
    
    ' Modify the refers to formula of the defined name "summa"
    oDefName.RefersTo = "=SUM(A1:B1)"
    
    ' Set descriptive text in cell A3
    oWorksheet.Range("A3").Value = "The name 'summa' refers to the formula from the cell C1."
    
    ' Set the formula description in cell A4
    oWorksheet.Range("A4").Value = "Formula: " & oDefName.RefersTo
End Sub
```

```js
// OnlyOffice JavaScript Code
// This example shows how to get a formula that the name is defined to refer to.
function defineAndManipulateNamedRange() {
    var oWorksheet = Api.GetActiveSheet();
    
    // Set values in cells A1 and B1
    oWorksheet.GetRange("A1").SetValue("1");
    oWorksheet.GetRange("B1").SetValue("2");
    
    // Set formula in cell C1
    oWorksheet.GetRange("C1").SetValue("=SUM(A1:B1)");
    
    // Add a defined name "summa" referring to the range A1:B1
    Api.AddDefName("summa", "Sheet1!$A$1:$B$1");
    
    // Get the defined name "summa"
    var oDefName = Api.GetDefName("summa");
    
    // Set the refers to formula of the defined name "summa"
    oDefName.SetRefersTo("=SUM(A1:B1)");
    
    // Set descriptive text in cell A3
    oWorksheet.GetRange("A3").SetValue("The name 'summa' refers to the formula from the cell C1.");
    
    // Set the formula description in cell A4
    oWorksheet.GetRange("A4").SetValue("Formula: " + oDefName.GetRefersTo());
}
```