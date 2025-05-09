# Range HasArray Property

## Business Description
True if the specified cell is part of an array formula. Read-only Variant.

## Behavior
Trueif the specified cell is part of an array formula. Read-onlyVariant.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
If ActiveCell.HasArray=True Then 
 MsgBox "The active cell is part of an array" 
End If
```