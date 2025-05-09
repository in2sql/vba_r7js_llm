# ErrorCheckingOptions InconsistentFormula Property

## Business Description
When set to True (default), Microsoft Excel identifies cells containing an inconsistent formula in a region. False disables the inconsistent formula check. Read/write Boolean.

## Behavior
When set toTrue(default), Microsoft Excel identifies cells containing an inconsistent formula in a region.Falsedisables the inconsistent formula check. Read/writeBoolean.

## Example Usage
```vba
Sub CheckFormula() 
 
 Application.ErrorCheckingOptions.InconsistentFormula= True 
 
 Range("A1:A3").Value = 1 
 Range("B1:B3").Value = 2 
 Range("C1:C3").Value = 3 
 
 Range("A4").Formula = "=SUM(A1:A3)" ' Consistent formula. 
 Range("B4").Formula = "=SUM(B1:B2)" ' Inconsistent formula. 
 Range("C4").Formula = "=SUM(C1:C3)" ' Consistent formula. 
 
End Sub
```