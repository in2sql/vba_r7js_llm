# PageSetup PrintTitleRows Property

## Business Description
Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro. Read/write String.

## Behavior
Returns or sets the rows that contain the cells to be repeated at the top of each page, as a string in A1-style notation in the language of the macro. Read/writeString.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.PrintTitleRows= ActiveSheet.Rows(3).Address 
ActiveSheet.PageSetup.PrintTitleColumns = _ 
 ActiveSheet.Columns("A:C").Address
```