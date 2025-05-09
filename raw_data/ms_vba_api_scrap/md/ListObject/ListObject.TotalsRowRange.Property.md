# ListObject TotalsRowRange Property

## Business Description
Returns a Range representing the Total row, if any, from a specified ListObject object. Read-only.

## Behavior
Returns aRangerepresenting the Total row, if any, from a specifiedListObjectobject. Read-only.

## Example Usage
```vba
Sub DisplayTotalsRowAddress() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet2") 
 Set objListObj = wrksht.ListObjects(1) 
 objListObj.ShowTotals = True 
 MsgBox objListObj.TotalsRowRange.Address 
End Sub
```