# VPageBreaks Add Method

## Business Description
Adds a vertical page break.

## Behavior
Adds a vertical page break.

## Example Usage
```vba
With Worksheets(1) 
 .HPageBreaks.Add .Range("F25") 
 .VPageBreaks.Add.Range("F25") 
End With
```