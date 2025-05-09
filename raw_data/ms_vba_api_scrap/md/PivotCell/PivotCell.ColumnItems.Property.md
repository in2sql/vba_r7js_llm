# PivotCell ColumnItems Property

## Business Description
Returns a PivotItemList collection that corresponds to the items on the column axis that represent the selected range.

## Behavior
Returns aPivotItemListcollection that corresponds to the items on the column axis that represent the selected range.

## Example Usage
```vba
Sub CheckColumnItems() 
 
 ' Determine if there is a match between the item and column field. 
 If Application.Range("B5").PivotCell.ColumnItems.Item(1) = "Inventory" Then 
 MsgBox "Item in B5 is a member of the 'Inventory' column field." 
 Else 
 MsgBox "Item in B5 is not a member of the 'Inventory' column field." 
 End If 
 
End Sub
```