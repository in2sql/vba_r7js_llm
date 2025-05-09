# HPageBreaks Add Method

## Business Description
Adds a horizontal page break.

## Behavior
Adds a horizontal page break.

## Example Usage
```vba
With Worksheets(1) 
 .HPageBreaks.Add.Range("F25") 
 .VPageBreaks.Add .Range("F25") 
End With
```