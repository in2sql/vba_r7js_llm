# Comment Author Property

## Business Description
Returns or sets the author of the comment. Read-only String.

## Behavior
Returns or sets the author of the comment. Read-onlyString.

## Example Usage
```vba
For Each c in ActiveSheet.Comments 
 If c.Author= "Jean Selva" Then c.Delete 
Next
```