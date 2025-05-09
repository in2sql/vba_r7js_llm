# Hyperlinks Object

## Business Description
Represents the collection of hyperlinks for a worksheet or range.

## Behavior
Represents the collection of hyperlinks for a worksheet or range.

## Example Usage
```vba
With Worksheets(1) 
 .Hyperlinks.Add .Range("E5"), "http://example.microsoft.com" 
End With
```