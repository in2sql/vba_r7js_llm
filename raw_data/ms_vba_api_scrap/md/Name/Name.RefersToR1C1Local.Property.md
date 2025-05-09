# Name RefersToR1C1Local Property

## Business Description
Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign. Read/write String.

## Behavior
Returns or sets the formula that the name refers to. This formula is in the language of the user, and it's in R1C1-style notation, beginning with an equal sign. Read/writeString.

## Example Usage
```vba
Set newSheet = ActiveWorkbook.Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.NameLocal 
 newSheet.Cells(i, 2).Value = "'" & nm.RefersToR1C1Locali = i + 1 
Next
```