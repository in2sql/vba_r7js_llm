# ErrorCheckingOptions EvaluateToError Property

## Business Description
When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain formulas evaluating to an error. False disables error checking for cells that evaluate to an error value. Read/write Boolean.

## Behavior
When set toTrue(default), Microsoft Excel identifies, with anAutoCorrect Optionsbutton, selected cells that contain formulas evaluating to an error.Falsedisables error checking for cells that evaluate to an error value. Read/writeBoolean.

## Example Usage
```vba
Sub CheckEvaluationError() 
 
 ' Simulate a divide-by-zero error. 
 Application.ErrorCheckingOptions.EvaluateToError= True 
 Range("A1").Value = 1 
 Range("A2").Value = 0 
 Range("A3").Formula = "=A1/A2" 
 
End Sub
```