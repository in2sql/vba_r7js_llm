# QueryTable BeforeRefresh Event

## Business Description
Occurs before any refreshes of the query table. This includes refreshes resulting from calling the Refresh method, from the user's actions in the product, and from opening the workbook containing the query table.

## Behavior
Occurs before any refreshes of the query table. This includes refreshes resulting from calling theRefreshmethod, from the user's actions in the product, and from opening the workbook containing the query table.

## Example Usage
```vba
Private Sub QueryTable_BeforeRefresh(Cancel As Boolean) 
 a = MsgBox("Refresh Now?", vbYesNoCancel) 
 If a = vbNo Then Cancel = True 
 MsgBox Cancel 
End Sub
```