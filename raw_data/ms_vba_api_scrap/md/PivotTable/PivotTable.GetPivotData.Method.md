# PivotTable GetPivotData Method

## Business Description
Returns a Range object with information about a data item in a PivotTable report.

## Behavior
Returns aRangeobject with information about a data item in a PivotTable report.

## Example Usage
```vba
Sub UseGetPivotData() 
 
 Dim rngTableItem As Range 
 
 ' Get PivotData for the quantity of chairs in the warehouse. 
 Set rngTableItem = ActiveCell. _ 
 PivotTable.GetPivotData("Quantity", "Warehouse", "Chairs") 
 
 MsgBox "The quantity of chairs in the warehouse is: " & rngTableItem.Value 
 
End Sub
```