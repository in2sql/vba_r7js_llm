# Range FindPrevious Method

## Business Description
Continues a search that was begun with the Find method. Finds the previous cell that matches those same conditions and returns a Range object that represents that cell. Doesn't affect the selection or the active cell.

## Behavior
Continues a search that was begun with theFindmethod. Finds the previous cell that matches those same conditions and returns aRangeobject that represents that cell. Doesn't affect the selection or the active cell.

## Example Usage
```vba
Sub FindTest() 
 Dim fc As Range 
 Set fc = Worksheets("Sheet1").Columns("B").Find(what:="Phoenix") 
 MsgBox "The first occurrence is in cell " & fc.Address 
 Set fc = Worksheets("Sheet1").Columns("B").FindNext(after:=fc) 
 MsgBox "The next occurrence is in cell " & fc.Address 
 Set fc = Worksheets("Sheet1").Columns("B").FindPrevious(after:=fc) 
 MsgBox "The previous occurrence is in cell " & fc.Address 
End Sub
```