# PageSetup TopMargin Property

## Business Description
Returns or sets the size of the top margin, in points. Read/write Double.

## Behavior
Returns or sets the size of the top margin, in points. Read/writeDouble.

## Example Usage
```vba
marginInches = ActiveSheet.PageSetup.TopMargin/ _ 
 Application.InchesToPoints(1) 
MsgBox "The current top margin is " & marginInches & " inches"
```