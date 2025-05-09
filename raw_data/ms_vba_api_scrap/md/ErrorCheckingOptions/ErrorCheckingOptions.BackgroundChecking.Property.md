# ErrorCheckingOptions BackgroundChecking Property

## Business Description
Alerts the user for all cells that violate enabled error-checking rules. When this property is set to True (default), the AutoCorrect Options button appears next to all cells that violate enabled errors. False disables background checking for errors.

## Behavior
Alerts the user for all cells that violate enabled error-checking rules. When this property is set toTrue(default), theAutoCorrect Optionsbutton appears next to all cells that violate enabled errors.Falsedisables background checking for errors. Read/writeBoolean.

## Example Usage
```vba
Sub CheckBackground() 
 
 ' Simulate an error by referring to empty cells. 
 Application.ErrorCheckingOptions.BackgroundChecking= True 
 Range("A1").Select 
 ActiveCell.Formula = "=A2/A3" 
 
End Sub
```