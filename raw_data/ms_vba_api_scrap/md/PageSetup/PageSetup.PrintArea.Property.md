# PageSetup PrintArea Property

## Business Description
Returns or sets the range to be printed, as a string using A1-style references in the language of the macro. Read/write String.

## Behavior
Returns or sets the range to be printed, as a string using A1-style references in the language of the macro. Read/writeString.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.PrintArea= _ 
 ActiveCell.CurrentRegion.Address
```