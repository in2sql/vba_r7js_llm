### This example recalculates all formulas in the active workbook.
### Этот пример пересчитывает все формулы в активной рабочей книге.

```vba
' Excel VBA code to recalculate all formulas in the active workbook

Sub RecalculateFormulas()
    Dim ws As Worksheet
    Dim rng As Range
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set values in B1 and C1
    ws.Range("B1").Value = 1
    ws.Range("C1").Value = 2
    
    ' Set formula in A1
    Set rng = ws.Range("A1")
    rng.Formula = "=SUM(B1:C1)"
    
    ' Set formula in E1
    Set rng = ws.Range("E1")
    rng.Formula = "=A1+1"
    
    ' Update value in B1
    ws.Range("B1").Value = 3
    
    ' Recalculate all formulas
    Application.CalculateFull
    
    ' Set value in A3 with a message
    ws.Range("A3").Value = "Formulas from cells A1 and E1 were recalculated with a new value from cell C1."
End Sub
```

```javascript
// This example recalculates all formulas in the active workbook.

function recalculateFormulas() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set values in B1 and C1
    oWorksheet.GetRange("B1").SetValue(1);
    oWorksheet.GetRange("C1").SetValue(2);
    
    // Set formula in A1
    var oRange = oWorksheet.GetRange("A1");
    oRange.SetValue("=SUM(B1:C1)");
    
    // Set formula in E1
    oRange = oWorksheet.GetRange("E1");
    oRange.SetValue("=A1+1");
    
    // Update value in B1
    oWorksheet.GetRange("B1").SetValue(3);
    
    // Recalculate all formulas
    Api.RecalculateAllFormulas();
    
    // Set value in A3 with a message
    oWorksheet.GetRange("A3").SetValue("Formulas from cells A1 and E1 were recalculated with a new value from cell C1.");
}

recalculateFormulas();
```