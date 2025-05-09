# ErrorCheckingOptions NumberAsText Property

## Business Description
When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, selected cells that contain numbers written as text. False disables error checking for numbers written as text. Read/write Boolean.

## Behavior
When set toTrue(default), Microsoft Excel identifies, with anAutoCorrect Optionsbutton, selected cells that contain numbers written as text.Falsedisables error checking for numbers written as text. Read/writeBoolean.

## Example Usage
```vba
Sub CheckNumberAsText() 
 
 ' Simulate an error by referencing a number stored as text. 
 Application.ErrorCheckingOptions.NumberAsText= True 
 Range("A1").Value = "'1" 
 
End Sub
```