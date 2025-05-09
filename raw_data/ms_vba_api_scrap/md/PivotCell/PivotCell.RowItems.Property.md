# PivotCell RowItems Property

## Business Description
Returns a PivotItemList collection that corresponds to the items on the category axis that represent the selected cell.

## Behavior
Returns aPivotItemListcollection that corresponds to the items on the category axis that represent the selected cell.

## Example Usage
```vba
Sub CheckRowItems() 
 
 ' Determine if there is a match between the item and row field. 
 If Application.Range("B5").PivotCell.RowItems.Item(1) = "Inventory" Then 
 MsgBox "Cell B5 is a member of the 'Inventory' row field. 
 Else 
 MsgBox "Cell B5 is not a member of the 'Inventory' row field. 
 End If 
 
End Sub
```