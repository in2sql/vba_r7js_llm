# CellFormat Font Property

## Business Description
Returns a Font object, allowing the user to set or return the search criteria based on the cell's font format.

## Behavior
Returns aFontobject, allowing the user to set or return the search criteria based on the cell's font format.

## Example Usage
```vba
Sub SearchCellFormat() 
 
 ' Set the search criteria for the font of the cell format. 
 Application.FindFormat.Font.ColorIndex = 3 
 
 ' Set the color index of the font for cell A5 to red. 
 Range("A5").Font.ColorIndex = 3 
 Range("A5").Formula = "Red font" 
 Range("A1").Select 
 MsgBox "Cell A5 has red font" 
 
 ' Find the cells based on the search criteria. 
 Cells.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _ 
 xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _ 
 , SearchFormat:=True).Activate 
 
 MsgBox "Microsoft Excel has found this cell matching the search criteria." 
 
End Sub
```