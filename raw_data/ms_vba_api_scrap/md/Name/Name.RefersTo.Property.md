# Name RefersTo Property

## Business Description
Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign. Read/write String.

## Behavior
Returns or sets the formula that the name is defined to refer to, in the language of the macro and in A1-style notation, beginning with an equal sign. Read/writeString.

## Example Usage
```vba
Set newSheet = Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.Name 
 newSheet.Cells(i, 2).Value = "'" & nm.RefersToi = i + 1 
Next 
newSheet.Columns("A:B").AutoFit
```