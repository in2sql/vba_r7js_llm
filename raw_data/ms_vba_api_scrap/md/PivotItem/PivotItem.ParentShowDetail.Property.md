# PivotItem ParentShowDetail Property

## Business Description
True if the specified item is showing because one of its parents is showing detail. False if the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped. Read-only Boolean.

## Behavior
Trueif the specified item is showing because one of its parents is showing detail.Falseif the specified item isn't showing because one of its parents is hiding detail. This property is available only if the item is grouped. Read-onlyBoolean.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Set pvtItem = ActiveCell.PivotItem 
If pvtItem.ParentShowDetail= True Then 
 MsgBox "Parent item is showing detail" 
End If
```