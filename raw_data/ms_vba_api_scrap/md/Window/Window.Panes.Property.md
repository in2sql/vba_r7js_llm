# Window Panes Property

## Business Description
Returns a Panes collection that represents all the panes in the specified window. Read-only.

## Behavior
Returns aPanescollection that represents all the panes in the specified window. Read-only.

## Example Usage
```vba
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
MsgBox "There are " & ActiveWindow.Panes.Count & _ 
 " panes in the active window"
```