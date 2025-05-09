# Range MergeArea Property

## Business Description
Returns a Range object that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell. Read-only Variant.

## Behavior
Returns aRangeobject that represents the merged range containing the specified cell. If the specified cell isn't in a merged range, this property returns the specified cell. Read-onlyVariant.

## Example Usage
```vba
Set ma = Range("a3").MergeAreaIf ma.Address = "$A$3" Then 
 MsgBox "not merged" 
Else 
 ma.Cells(1, 1).Value = "42" 
End If
```