# Error Object

## Business Description
Represents a spreadsheet error for a range.

## Behavior
Represents a spreadsheet error for a range.

## Example Usage
```vba
Sub CheckEmptyCells() 
 
 Dim rngFormula As Range 
 Set rngFormula = Application.Range("A1") 
 
 ' Place a formula referencing empty cells. 
 Range("A1").Formula = "=A2+A3" 
 Application.ErrorCheckingOptions.EmptyCellReferences = True 
 
 ' Perform check to see if EmptyCellReferences check is on. 
 If rngFormula.Errors.Item(xlEmptyCellReferences).Value = True Then 
 MsgBox "The empty cell references error checking feature is enabled." 
 Else 
 MsgBox "The empty cell references error checking feature is not on." 
 End If 
 
End Sub
```