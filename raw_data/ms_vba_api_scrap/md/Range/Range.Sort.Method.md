# Range Sort Method

## Business Description
Sorts a range of values.

## Behavior
Sorts a range of values.

## Example Usage
```vba
Sub ColorSort()
   'Set up your variables and turn off screen updating.
   Dim iCounter As Integer
   Application.ScreenUpdating = False
   
   'For each cell in column A, go through and place the color index value of the cell in column C.
   For iCounter = 2 To 55
      Cells(iCounter, 3) = _
         Cells(iCounter, 1).Interior.ColorIndex
   Next iCounter
   
   'Sort the rows based on the data in column C
   Range("C1") = "Index"
   Columns("A:C").Sort key1:=Range("C2"), _
      order1:=xlAscending, header:=xlYes
   
   'Clear out the temporary sorting value in column C, and turn screen updating back on.
   Columns(3).ClearContents
   Application.ScreenUpdating = True
End Sub
```