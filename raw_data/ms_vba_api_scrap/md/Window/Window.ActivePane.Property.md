# Window ActivePane Property

## Business Description
Returns a Pane object that represents the active pane in the window. Read-only.

## Behavior
Returns aPaneobject that represents the active pane in the window. Read-only.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Activate 
If not ActiveWindow.FreezePanes Then 
 With ActiveWindow 
 i = .ActivePane.Index 
 If i = .Panes.Count Then 
 .Panes(1).Activate 
 Else 
 .Panes(i+1).Activate 
 End If 
 End With 
End If
```