# How to: Refer to Cells and Ranges by Using A1 Notation

## Business Description
You can refer to a cell or range of cells in the A1 reference style by using the Range property. The following subroutine changes the format of cells A1:D5 to bold.

## Behavior
You can refer to a cell or range of cells in the A1 reference style by using theRangeproperty. The following subroutine changes the format of cells A1:D5 to bold.

## Example Usage
```vba
Sub FormatRange() 
 Workbooks("Book1").Sheets("Sheet1").Range("A1:D5") _ 
 .Font.Bold = True 
End Sub
```