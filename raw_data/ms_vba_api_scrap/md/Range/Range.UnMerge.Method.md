# Range UnMerge Method

## Business Description
Separates a merged area into individual cells.

## Behavior
Separates a merged area into individual cells.

## Example Usage
```vba
With Range("a3") 
 If .MergeCells Then 
 .MergeArea.UnMergeElse 
 MsgBox "not merged" 
 End If 
End With
```