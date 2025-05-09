# CellFormat Object

## Business Description
Represents the search criteria for the cell format.

## Behavior
Represents the search criteria for the cell format.

## Example Usage
```vba
Sub ChangeCellFormat() 
 
 ' Set the interior of cell A1 to yellow. 
 Range("A1").Select 
 Selection.Interior.ColorIndex = 36 
 MsgBox "The cell format for cell A1 is a yellow interior." 
 
 ' Set the CellFormat object to replace yellow with green. 
 With Application 
 .FindFormat.Interior.ColorIndex = 36 
 .ReplaceFormat.Interior.ColorIndex = 35 
 End With 
 
 ' Find and replace cell A1's yellow interior with green. 
 ActiveCell.Replace What:="", Replacement:="", LookAt:=xlPart, _ 
 SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=True, _ 
 ReplaceFormat:=True 
 MsgBox "The cell format for cell A1 is replaced with a green interior." 
 
End Sub
```