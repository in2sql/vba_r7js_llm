# ErrorCheckingOptions OmittedCells Property

## Business Description
When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included. False disables error checking for omitted cells.

## Behavior
When set toTrue(default), Microsoft Excel identifies, with an AutoCorrect Options button, the selected cells that contain formulas referring to a range that omits adjacent cells that could be included.Falsedisables error checking for omitted cells. Read/writeBoolean.

## Example Usage
```vba
Sub CheckOmittedCells() 
 
 Application.ErrorCheckingOptions.OmittedCells= True 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Value = 3 
 Range("A4").Formula = "=Sum(A1:A2)" 
 
End Sub
```