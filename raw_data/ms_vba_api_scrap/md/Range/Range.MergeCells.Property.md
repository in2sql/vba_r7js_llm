# Range MergeCells Property

## Business Description
True if the range contains merged cells. Read/write Variant.

## Behavior
Trueif the range contains merged cells. Read/writeVariant.

## Example Usage
```vba
Set ma = Range("a3").MergeAreaIf Range("a3").MergeCellsThen 
 ma.Cells(1, 1).Value = "42" 
End If
```