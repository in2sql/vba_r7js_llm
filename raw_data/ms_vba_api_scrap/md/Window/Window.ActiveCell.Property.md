# Window ActiveCell Property

## Business Description
Returns a Range object that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.

## Behavior
Returns aRangeobject that represents the active cell in the active window (the window on top) or in the specified window. If the window isn't displaying a worksheet, this property fails. Read-only.

## Example Usage
```vba
ActiveCell 
Application.ActiveCell 
ActiveWindow.ActiveCell 
Application.ActiveWindow.ActiveCell
```