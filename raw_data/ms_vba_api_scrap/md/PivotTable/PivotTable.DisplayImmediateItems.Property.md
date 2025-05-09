# PivotTable DisplayImmediateItems Property

## Business Description
Returns or sets a Boolean that indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty.

## Behavior
Returns or sets aBooleanthat indicates whether items in the row and column areas are visible when the data area of the PivotTable is empty. Set this property toFalseto hide the items in the row and column areas when the data area of the PivotTable is empty. The default value isTrue.

## Example Usage
```vba
Sub CheckItemsDisplayed() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine how the PivotTable was created. 
 If pvtTable.DisplayImmediateItems= True Then 
 MsgBox "Fields have been added to the row or column areas for the PivotTable report." 
 Else 
 MsgBox "The PivotTable was created by using object-model calls." 
 End If 
 
End Sub
```