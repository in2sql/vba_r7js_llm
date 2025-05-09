# Name RefersToR1C1 Property

## Business Description
Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign. Read/write String.

## Behavior
Returns or sets the formula that the name refers to. The formula is in the language of the macro, and it's in R1C1-style notation, beginning with an equal sign. Read/writeString.

## Example Usage
```vba
Set newSheet = ActiveWorkbook.Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.Name 
 newSheet.Cells(i, 2).Value = "'" & nm.RefersToR1C1i = i + 1 
Next
```