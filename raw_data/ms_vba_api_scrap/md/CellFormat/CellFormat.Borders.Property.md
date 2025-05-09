# CellFormat Borders Property

## Business Description
Returns or sets a Borders collection that represents the search criteria based on the cell's border format.

## Behavior
Returns or sets aBorderscollection that represents the search criteria based on the cell's border format.

## Example Usage
```vba
Sub SearchCellFormat() 
 
 ' Set the search criteria for the border of the cell format. 
 With Application.FindFormat.Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThick 
 End With 
 
 ' Create a continuous thick bottom-edge border for cell A5. 
 Range("A5").Select 
 With Selection.Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThick 
 End With 
 Range("A1").Select 
 MsgBox "Cell A5 has a continuous thick bottom-edge border" 
 
 ' Find the cells based on the search criteria. 
 Cells.Find(What:="", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _ 
 xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _ 
 , SearchFormat:=True).Activate 
 MsgBox "Microsoft Excel has found this cell matching the search criteria." 
 
End Sub
```