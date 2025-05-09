# Window WindowNumber Property

## Business Description
Returns the window number. For example, a window named "Book1.xls:2" has 2 as its window number. Most windows have the window number 1. Read-only Long.

## Behavior
Returns the window number. For example, a window named "Book1.xls:2" has 2 as its window number. Most windows have the window number 1. Read-onlyLong.

## Example Usage
```vba
ActiveWindow.NewWindow 
MsgBox ActiveWindow.WindowNumber
```