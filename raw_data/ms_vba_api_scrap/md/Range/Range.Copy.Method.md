# Range Copy Method

## Business Description
Copies the range to the specified range or to the Clipboard.

## Behavior
Copies the range to the specified range or to the Clipboard.

## Example Usage
```vba
Worksheets("Sheet1").Range("A1:D4").Copy_ 
    destination:=Worksheets("Sheet2").Range("E5")
```