# How to: Select a Range

## Business Description
These examples show how to select the used range, which includes formatted cells that do not contain data, and how to select a data range, which includes cells that contains actual data.

## Behavior
These examples show how to select the used range, which includes formatted cells that do not contain data, and how to select a data range, which includes cells that contains actual data.

## Example Usage
```vba
Sub SelectUsedRange()
    ActiveSheet.UsedRange.Select
    MsgBox "The used range address is " & ActiveSheet.UsedRange.Address(0, 0) & ".", 64, "Used range address:"
End Sub
```