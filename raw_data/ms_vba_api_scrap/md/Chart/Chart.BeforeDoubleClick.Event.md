# Chart BeforeDoubleClick Event

## Business Description
Occurs when a chart element is double-clicked, before the default double-click action.

## Behavior
Occurs when a chart element is double-clicked, before the default double-click action.

## Example Usage
```vba
Private Sub Chart_BeforeDoubleClick(ByVal ElementID As Long, _ 
 ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean) 
 
 If ElementID = xlFloor Then 
 Cancel = True 
 MsgBox "Chart formatting for this item is restricted." 
 End If 
 
End Sub
```