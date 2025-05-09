### Description / Описание

**English:** This code recalculates all formulas in the active workbook by setting values to specific cells and updating formulas accordingly.

**Russian:** Этот код пересчитывает все формулы в активной рабочей книге, устанавливая значения в определенные ячейки и обновляя формулы соответственно.

---

#### VBA Code

```vba
' This code recalculates all formulas in the active workbook
Sub RecalculateFormulas()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value to cell B1
    oWorksheet.Range("B1").Value = 1
    
    ' Set value to cell C1
    oWorksheet.Range("C1").Value = 2
    
    ' Set formula in cell A1
    Set oRange = oWorksheet.Range("A1")
    oRange.Formula = "=SUM(B1:C1)"
    
    ' Set formula in cell E1
    Set oRange = oWorksheet.Range("E1")
    oRange.Formula = "=A1+1"
    
    ' Update value in cell B1
    oWorksheet.Range("B1").Value = 3
    
    ' Recalculate all formulas
    Application.CalculateFull
    
    ' Set message in cell A3
    oWorksheet.Range("A3").Value = "Formulas from cells A1 and E1 were recalculated with a new value from cell C1."
End Sub
```

---

#### OnlyOffice JavaScript Code

```javascript
// This example recalculates all formulas in the active workbook.

function RecalculateFormulas() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value to cell B1
    oWorksheet.GetRange("B1").SetValue(1);
    
    // Set value to cell C1
    oWorksheet.GetRange("C1").SetValue(2);
    
    // Set formula in cell A1
    var oRange = oWorksheet.GetRange("A1");
    oRange.SetValue("=SUM(B1:C1)");
    
    // Set formula in cell E1
    oRange = oWorksheet.GetRange("E1");
    oRange.SetValue("=A1+1");
    
    // Update value in cell B1
    oWorksheet.GetRange("B1").SetValue(3);
    
    // Recalculate all formulas
    Api.RecalculateAllFormulas();
    
    // Set message in cell A3
    oWorksheet.GetRange("A3").SetValue("Formulas from cells A1 and E1 were recalculated with a new value from cell C1.");
}

// Execute the function
RecalculateFormulas();
```