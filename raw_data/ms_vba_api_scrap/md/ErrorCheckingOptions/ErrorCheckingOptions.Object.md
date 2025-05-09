# ErrorCheckingOptions Object

## Business Description
Represents the error-checking options for an application.

## Behavior
Represents the error-checking options for an application.

## Example Usage
```vba
Sub CheckTextDates() 
 
 Dim rngFormula As Range 
 Set rngFormula = Application.Range("A1") 
 
 Range("A1").Formula = "'April 23, 00" 
 Application.ErrorCheckingOptions.TextDate = True 
 
 ' Perform check to see if 2 digit year TextDate check is on. 
 If rngFormula.Errors.Item(xlTextDate).Value = True Then 
 MsgBox "The text date error checking feature is enabled." 
 Else 
 MsgBox "The text date error checking feature is not on." 
 End If 
 
End Sub
```