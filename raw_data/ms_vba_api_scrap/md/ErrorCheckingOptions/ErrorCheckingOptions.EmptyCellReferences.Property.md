# ErrorCheckingOptions EmptyCellReferences Property

## Business Description
When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells containing formulas that refer to empty cells. False disables empty cell reference checking. Read/write Boolean.

## Behavior
When set toTrue(default), Microsoft Excel identifies, with anAutoCorrect Optionsbutton, selected cells containing formulas that refer to empty cells.Falsedisables empty cell reference checking. Read/writeBoolean.

## Example Usage
```vba
Sub CheckEmptyCells() 
 
 Application.ErrorCheckingOptions.EmptyCellReferences= True 
 Range("A1").Formula = "=A2+A3" 
 
End Sub
```