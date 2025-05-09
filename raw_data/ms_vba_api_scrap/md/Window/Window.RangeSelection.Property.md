# Window RangeSelection Property

## Business Description
Returns a Range object that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. Read-only.

## Behavior
Returns aRangeobject that represents the selected cells on the worksheet in the specified window even if a graphic object is active or selected on the worksheet. Read-only.

## Example Usage
```vba
MsgBox ActiveWindow.RangeSelection.Address
```