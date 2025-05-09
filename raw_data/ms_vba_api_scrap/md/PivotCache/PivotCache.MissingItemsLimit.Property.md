# PivotCache MissingItemsLimit Property

## Business Description
Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/write XlPivotTableMissingItems.

## Behavior
Returns or sets the maximum quantity of unique items per PivotTable field that are retained even when they have no supporting data in the cache records. Read/writeXlPivotTableMissingItems.

## Example Usage
```vba
Sub CheckMissingItemsList() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the maximum number of unique items allowed per PivotField and notify the user. 
 Select Case pvtCache.MissingItemsLimitCase xlMissingItemsDefault 
 MsgBox "The default value of unique items per PivotField is allowed." 
 Case xlMissingItemsMax 
 MsgBox "The maximum value of unique items per PivotField is allowed." 
 Case xlMissingItemsNone 
 MsgBox "No unique items per PivotField are allowed." 
 End Select 
 
End Sub
```