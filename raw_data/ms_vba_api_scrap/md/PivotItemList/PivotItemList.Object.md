# PivotItemList Object

## Business Description
A collection of all the PivotItem objects in the specified PivotTable.

## Behavior
A collection of all thePivotItemobjects in the specified PivotTable.

## Example Usage
```vba
Sub CheckPivotItemList() 
 
 ' Identify contents associated with PivotItemList. 
 MsgBox "Contents associated with cell B5: " & _ 
 Application.Range("B5").PivotCell.RowItems.Item(1) 
 
End Sub
```