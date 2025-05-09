# CellFormat Interior Property

## Business Description
Returns an Interior object allowing the user to set or return the search criteria based on the cell's interior format.

## Behavior
Returns anInteriorobject allowing the user to set or return the search criteria based on the cell's interior format.

## Example Usage
```vba
Sub SearchCellFormat() 
 
 ' Set the search criteria for the interior of the cell format. 
 With Application.FindFormat.Interior.ColorIndex = 6 
 .Pattern = xlSolid 
 .PatternColorIndex = xlAutomatic 
 End With 
 
 ' Create a yellow interior for cell A5. 
 Range("A5").Select 
 With Selection.Interior 
 .ColorIndex = 6 
 .Pattern = xlSolid 
 .PatternColorIndex = xlAutomatic 
 End With 
 Range("A1").Select 
 MsgBox "Cell A5 has a yellow interior." 
 
 ' Find the cells based on the search criteria. 
 Cells.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _ 
 xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _ 
 , SearchFormat:=True).Activate 
 MsgBox "Microsoft Excel has found this cell matching the search criteria." 
 
End Sub
```