# TextFrame2 HasText Property

## Business Description
Returns whether the specified text frame has text. Read-only MsoTriStatehttp://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515(Office.15).aspx.

## Behavior
Returns whether the specified text frame has text. Read-onlyMsoTriState.

## Example Usage
```vba
With ActiveSheet.Shapes(1).TextFrame2 
If .HasText Then 
.TextRange2.Font.Name = "Arial"
```