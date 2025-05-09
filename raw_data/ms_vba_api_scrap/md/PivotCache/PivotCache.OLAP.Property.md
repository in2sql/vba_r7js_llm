# PivotCache OLAP Property

## Business Description
Returns True if the PivotTable cache is connected to an Online Analytical Processing (OLAP) server. Read-onlyBoolean.

## Behavior
ReturnsTrueif the PivotTable cache is connected to an Online Analytical Processing (OLAP) server. Read-onlyBoolean.

## Example Usage
```vba
Sub CheckPivotCache() 
 
 ' Determine if PivotCache has OLAP connection. 
 If Application.ActiveWorkbook.PivotCaches.Item(1).OLAP= True Then 
 MsgBox "The PivotCache is connected to an OLAP server" 
 Else 
 MsgBox "The PivotCache is not connected to an OLAP server." 
 End If 
 
End Sub
```