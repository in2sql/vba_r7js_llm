# ErrorCheckingOptions UnlockedFormulaCells Property

## Business Description
When set to True (default), Microsoft Excel identifies selected cells that are unlocked and contain a formula. False disables error checking for unlocked cells that contain formulas. Read/write Boolean.

## Behavior
When set toTrue(default), Microsoft Excel identifies selected cells that are unlocked and contain a formula.Falsedisables error checking for unlocked cells that contain formulas. Read/writeBoolean.

## Example Usage
```vba
Sub CheckUnlockedCell() 
 
 Application.ErrorCheckingOptions.UnlockedFormulaCells= True 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Formula = "=A1+A2" 
 Range("A3").Locked = False 
 
End Sub
```