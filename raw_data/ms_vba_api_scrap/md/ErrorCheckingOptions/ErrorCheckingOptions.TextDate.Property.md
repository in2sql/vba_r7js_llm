# ErrorCheckingOptions TextDate Property

## Business Description
When set to True (default), Microsoft Excel identifies, with an AutoCorrect Options button, cells that contain a text date with a two-digit year. False disables error checking for cells containing a text date with a two-digit year. Read/write Boolean.

## Behavior
When set toTrue(default), Microsoft Excel identifies, with anAutoCorrect Optionsbutton, cells that contain a text date with a two-digit year.Falsedisables error checking for cells containing a text date with a two-digit year. Read/writeBoolean.

## Example Usage
```vba
Sub CheckTextDate() 
 
 ' Simulate an error by referencing a text date with a two-digit year. 
 Application.ErrorCheckingOptions.TextDate= True 
 Range("A1").Formula = "'April 23, 00" 
 
End Sub
```